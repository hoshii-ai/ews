package ews

import (
	"bytes"
	"context"
	"crypto/tls"
	"encoding/xml"
	"fmt"
	"io"
	"net/http"
	"net/http/httputil"

	ntlmssp "github.com/Azure/go-ntlmssp"
)

const (
	// DefaultServerVersion is the RequestServerVersion sent when Config.ServerVersion is empty.
	DefaultServerVersion = "Exchange2013_SP1"

	soapEnd = `
</soap:Body></soap:Envelope>`
)

func soapStart(serverVersion string) string {
	var v bytes.Buffer
	_ = xml.EscapeText(&v, []byte(serverVersion))
	return `<?xml version="1.0" encoding="utf-8" ?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
		xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  		<soap:Header>
    		<RequestServerVersion xmlns="http://schemas.microsoft.com/exchange/services/2006/types" Version="` + v.String() + `" />
  		</soap:Header>
  		<soap:Body>
`
}

type Config struct {
	Dump    bool
	NTLM    bool
	SkipTLS bool

	// ServerVersion is the RequestServerVersion value, e.g. Exchange2010_SP1,
	// Exchange2013_SP1, Exchange2016. Defaults to DefaultServerVersion.
	ServerVersion string

	// HTTPClient, if set, is used as-is for all requests: NTLM and SkipTLS are
	// ignored (configure them on its Transport). Redirects are not suppressed.
	// Its Timeout must be 0: a non-zero Timeout kills long-lived GetStreamingEvents
	// responses. Bound streams with the request context instead.
	HTTPClient *http.Client

	// Transport, if set, replaces the default transport underneath the client.
	// NTLM, when enabled, wraps it. SkipTLS is ignored (configure it on the transport).
	Transport http.RoundTripper
}

type Client interface {
	SendAndReceive(body []byte) ([]byte, error)
	GetEWSAddr() string
	GetUsername() string
}

// RequestOption customises a single request.
type RequestOption func(*http.Request)

// WithHeader sets an extra HTTP header on the request.
func WithHeader(key, value string) RequestOption {
	return func(r *http.Request) { r.Header.Set(key, value) }
}

// WithAnchorMailbox sets X-AnchorMailbox to the given mailbox SMTP address.
func WithAnchorMailbox(smtp string) RequestOption {
	return WithHeader("X-AnchorMailbox", smtp)
}

// ContextClient is a Client that supports contexts, per-request headers and
// streaming responses. Clients from NewClient/NewContextClient implement it.
type ContextClient interface {
	Client
	// SendAndReceiveContext sends body wrapped in a SOAP envelope and returns the whole response.
	SendAndReceiveContext(ctx context.Context, body []byte, opts ...RequestOption) ([]byte, error)
	// OpenStream sends body and returns the open response body without reading it.
	// Caller must Close it. Cancelling ctx aborts the read.
	OpenStream(ctx context.Context, body []byte, opts ...RequestOption) (io.ReadCloser, error)
}

type client struct {
	EWSAddr   string
	Username  string
	Password  string
	config    *Config
	http      *http.Client
	soapStart []byte
}

func (c *client) GetEWSAddr() string {
	return c.EWSAddr
}

func (c *client) GetUsername() string {
	return c.Username
}

// NewClient builds a Client. The underlying http.Client (and so its
// connections and NTLM state) is created once and reused for every request.
func NewClient(ewsAddr, username, password string, config *Config) Client {
	return NewContextClient(ewsAddr, username, password, config)
}

// NewContextClient is like NewClient but returns the richer ContextClient.
func NewContextClient(ewsAddr, username, password string, config *Config) ContextClient {
	if config == nil {
		config = &Config{}
	}
	version := config.ServerVersion
	if version == "" {
		version = DefaultServerVersion
	}
	return &client{
		EWSAddr:   ewsAddr,
		Username:  username,
		Password:  password,
		config:    config,
		http:      buildHTTPClient(config),
		soapStart: []byte(soapStart(version)),
	}
}

func buildHTTPClient(config *Config) *http.Client {
	if config.HTTPClient != nil {
		return config.HTTPClient
	}
	rt := config.Transport
	if rt == nil {
		t := http.DefaultTransport.(*http.Transport).Clone()
		if config.SkipTLS {
			t.TLSClientConfig = &tls.Config{InsecureSkipVerify: true}
		}
		rt = t
	}
	if config.NTLM {
		rt = ntlmssp.Negotiator{RoundTripper: rt}
	}
	return &http.Client{
		Transport: rt,
		CheckRedirect: func(req *http.Request, via []*http.Request) error {
			return http.ErrUseLastResponse
		},
	}
}

func (c *client) SendAndReceive(body []byte) ([]byte, error) {
	return c.SendAndReceiveContext(context.Background(), body)
}

func (c *client) SendAndReceiveContext(ctx context.Context, body []byte, opts ...RequestOption) ([]byte, error) {
	resp, err := c.do(ctx, body, opts, true)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	return io.ReadAll(resp.Body)
}

func (c *client) OpenStream(ctx context.Context, body []byte, opts ...RequestOption) (io.ReadCloser, error) {
	// response is not dumped: it would consume the stream
	resp, err := c.do(ctx, body, opts, false)
	if err != nil {
		return nil, err
	}
	return resp.Body, nil
}

// do sends the request and returns a 200 response. Non-200 responses are
// turned into errors (401 matches ErrUnauthorized) and their body closed.
func (c *client) do(ctx context.Context, body []byte, opts []RequestOption, dumpResp bool) (*http.Response, error) {
	bb := make([]byte, 0, len(c.soapStart)+len(body)+len(soapEnd))
	bb = append(bb, c.soapStart...)
	bb = append(bb, body...)
	bb = append(bb, soapEnd...)

	req, err := http.NewRequestWithContext(ctx, "POST", c.EWSAddr, bytes.NewReader(bb))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Content-Type", "text/xml")
	for _, o := range opts {
		o(req)
	}
	logRequest(c, req)
	req.SetBasicAuth(c.Username, c.Password)

	resp, err := c.http.Do(req)
	if err != nil {
		return nil, err
	}
	if dumpResp {
		logResponse(c, resp)
	}

	if resp.StatusCode != http.StatusOK {
		defer resp.Body.Close()
		if resp.StatusCode == http.StatusUnauthorized {
			return nil, &HTTPError{Status: resp.Status, StatusCode: resp.StatusCode}
		}
		return nil, NewError(resp)
	}
	return resp, nil
}

func logRequest(c *client, req *http.Request) {
	if c.config != nil && c.config.Dump {
		dump, err := httputil.DumpRequestOut(req, true)
		if err != nil {
			fmt.Println(err)
		}
		fmt.Printf("Request:\n%v\n----\n", string(dump))
	}
}

func logResponse(c *client, resp *http.Response) {
	if c.config != nil && c.config.Dump {
		dump, err := httputil.DumpResponse(resp, true)
		if err != nil {
			fmt.Println(err)
		}
		fmt.Printf("Response:\n%v\n----\n", string(dump))
	}
}
