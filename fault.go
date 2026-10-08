package ews

import (
	"encoding/xml"
	"errors"
	"io/ioutil"
	"net/http"
	"strings"
)

func NewError(resp *http.Response) error {
	soap, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		return err
	}
	fault, _ := parseSoapFault(string(soap))
	if fault == nil {
		return &HTTPError{Status: resp.Status, StatusCode: resp.StatusCode}
	}
	return &SoapError{Fault: fault}
}

type SoapError struct {
	Fault *Fault
}

func (s SoapError) Error() string {
	return s.Fault.Faultstring
}

// Unwrap exposes the EWS ResponseCode carried in the fault detail, if any, as
// a *ResponseError (e.g. ErrorServerBusy with BackOffMilliseconds).
func (s SoapError) Unwrap() error {
	if s.Fault == nil || s.Fault.Detail.ResponseCode == "" {
		return nil
	}
	d := s.Fault.Detail
	return &ResponseError{
		Class:               ResponseClassError,
		Code:                d.ResponseCode,
		MessageText:         d.Message,
		BackOffMilliseconds: backOff(d.MessageXml.Values),
	}
}

// ErrUnauthorized matches (errors.Is) an HTTP 401 response. Callers must not
// retry it: repeated bad credentials can lock the AD account.
var ErrUnauthorized = errors.New("ews: unauthorized (HTTP 401)")

type HTTPError struct {
	Status     string
	StatusCode int
}

func (s HTTPError) Is(target error) bool {
	return target == ErrUnauthorized && s.StatusCode == http.StatusUnauthorized
}

func (s HTTPError) Error() string {
	return s.Status
}

type envelop struct {
	XMLName struct{} `xml:"Envelope"`
	Body    body     `xml:"Body"`
}
type body struct {
	Fault Fault `xml:"Fault"`
}

type Fault struct {
	Faultcode   string `xml:"faultcode"`
	Faultstring string `xml:"faultstring"`
	Detail      detail `xml:"detail"`
}

type detail struct {
	ResponseCode string          `xml:"ResponseCode"`
	Message      string          `xml:"Message"`
	MessageXml   faultMessageXml `xml:"MessageXml"`
}

type faultMessageXml struct {
	LineNumber   string            `xml:"LineNumber"`
	LinePosition string            `xml:"LinePosition"`
	Violation    string            `xml:"Violation"`
	Values       []MessageXmlValue `xml:"Value"`
}

func parseSoapFault(soapMessage string) (*Fault, error) {
	var e envelop
	err := xml.Unmarshal([]byte(soapMessage), &e)
	if err != nil {
		return nil, err
	}
	if len(strings.TrimSpace(e.Body.Fault.Faultcode)) == 0 &&
		len(strings.TrimSpace(e.Body.Fault.Faultstring)) == 0 {
		return nil, nil
	}
	return &e.Body.Fault, nil
}
