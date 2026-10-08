package ews

import (
	"context"
	"errors"
	"io"
	"net"
	"net/http"
	"net/http/httptest"
	"os"
	"strings"
	"sync"
	"sync/atomic"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func fixture(t *testing.T, name string) string {
	t.Helper()
	b, err := os.ReadFile("testdata/" + name)
	require.NoError(t, err)
	return string(b)
}

// serve replies to each request with the next response (status, body); the last repeats.
type reply struct {
	status int
	body   string
}

func newServer(t *testing.T, replies ...reply) (*httptest.Server, *[]*http.Request, *[]string) {
	t.Helper()
	var mu sync.Mutex
	var reqs []*http.Request
	var bodies []string
	var n int32
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		b, _ := io.ReadAll(r.Body)
		mu.Lock()
		reqs = append(reqs, r)
		bodies = append(bodies, string(b))
		mu.Unlock()
		i := int(atomic.AddInt32(&n, 1)) - 1
		if i >= len(replies) {
			i = len(replies) - 1
		}
		w.WriteHeader(replies[i].status)
		_, _ = w.Write([]byte(replies[i].body))
	}))
	t.Cleanup(srv.Close)
	return srv, &reqs, &bodies
}

func TestServerVersion(t *testing.T) {
	srv, _, bodies := newServer(t, reply{200, fixture(t, "get_folder_inbox.xml")})
	ctx := context.Background()

	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := GetFolder(ctx, c, FolderRef{DistinguishedId: "inbox"})
	require.NoError(t, err)
	assert.Contains(t, (*bodies)[0], `Version="Exchange2013_SP1"`)

	c = NewContextClient(srv.URL, "u", "p", &Config{ServerVersion: "Exchange2010_SP1"})
	_, err = GetFolder(ctx, c, FolderRef{DistinguishedId: "inbox"})
	require.NoError(t, err)
	assert.Contains(t, (*bodies)[1], `Version="Exchange2010_SP1"`)
}

func TestLegacySendAndReceive(t *testing.T) {
	srv, reqs, _ := newServer(t, reply{200, "<ok/>"})
	c := NewClient(srv.URL, "u", "p", &Config{})
	out, err := c.SendAndReceive([]byte("<x/>"))
	require.NoError(t, err)
	assert.Equal(t, "<ok/>", string(out))
	u, p, _ := (*reqs)[0].BasicAuth()
	assert.Equal(t, "u", u)
	assert.Equal(t, "p", p)
	assert.Equal(t, "text/xml", (*reqs)[0].Header.Get("Content-Type"))
}

func TestGetFolderAndHeaders(t *testing.T) {
	srv, reqs, bodies := newServer(t, reply{200, fixture(t, "get_folder_inbox.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)

	res, err := GetFolder(context.Background(), c, FolderRef{DistinguishedId: "inbox"}, WithAnchorMailbox("a@b.com"))
	require.NoError(t, err)
	assert.Equal(t, "AAMkInbox==", res.FolderId)
	assert.Equal(t, "AQAAABYAAAInbox", res.ChangeKey)
	assert.Equal(t, "15", res.ServerVersionInfo.MajorVersion)
	assert.Equal(t, "1", res.ServerVersionInfo.MinorVersion)
	assert.Equal(t, "2507", res.ServerVersionInfo.MajorBuildNumber)
	assert.Equal(t, "39", res.ServerVersionInfo.MinorBuildNumber)
	assert.Equal(t, "V2017_07_11", res.ServerVersionInfo.Version)
	assert.Equal(t, "a@b.com", (*reqs)[0].Header.Get("X-AnchorMailbox"))
	assert.Contains(t, (*bodies)[0], `DistinguishedFolderId`)
	assert.Contains(t, (*bodies)[0], `IdOnly`)
}

func TestUnauthorizedNotRetriedAndSentinel(t *testing.T) {
	srv, reqs, _ := newServer(t, reply{401, fixture(t, "unauthorized_401.html")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := GetFolder(context.Background(), c, FolderRef{DistinguishedId: "inbox"})
	require.Error(t, err)
	assert.True(t, errors.Is(err, ErrUnauthorized))
	assert.Len(t, *reqs, 1)

	_, err = c.SendAndReceive([]byte("<x/>"))
	assert.True(t, errors.Is(err, ErrUnauthorized))
}

func TestServerBusyResponseError(t *testing.T) {
	srv, _, _ := newServer(t, reply{200, fixture(t, "error_server_busy.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := Subscribe(context.Background(), c, []FolderRef{{DistinguishedId: "inbox"}}, []EventType{EventNewMail})
	var re *ResponseError
	require.True(t, errors.As(err, &re))
	assert.Equal(t, "ErrorServerBusy", re.Code)
	assert.Equal(t, 45000, re.BackOffMilliseconds)
	assert.Contains(t, re.MessageText, "Try again later")
	assert.True(t, re.IsServerBusy())
}

func TestServerBusySoapFault(t *testing.T) {
	srv, _, _ := newServer(t, reply{503, fixture(t, "fault_server_busy.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := GetFolder(context.Background(), c, FolderRef{DistinguishedId: "inbox"})
	var se *SoapError
	require.True(t, errors.As(err, &se), "SoapError still returned")
	var re *ResponseError
	require.True(t, errors.As(err, &re))
	assert.Equal(t, "ErrorServerBusy", re.Code)
	assert.Equal(t, 30000, re.BackOffMilliseconds)
}

func TestSubscribeUnsubscribe(t *testing.T) {
	srv, _, bodies := newServer(t, reply{200, fixture(t, "subscribe_ok.xml")}, reply{200, fixture(t, "unsubscribe_ok.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	ctx := context.Background()

	id, err := Subscribe(ctx, c,
		[]FolderRef{{DistinguishedId: "inbox"}, {Id: "FID", ChangeKey: "CK"}},
		[]EventType{EventNewMail, EventCreated, EventMoved})
	require.NoError(t, err)
	assert.Equal(t, "SUB123==", id)
	b := (*bodies)[0]
	assert.Contains(t, b, "StreamingSubscriptionRequest")
	assert.Contains(t, b, `Id="inbox"`)
	assert.Contains(t, b, `Id="FID" ChangeKey="CK"`)
	assert.Contains(t, b, "NewMailEvent")
	assert.Contains(t, b, "MovedEvent")
	assert.Less(t, strings.Index(b, "FolderIds"), strings.Index(b, "EventTypes"), "schema order")

	require.NoError(t, Unsubscribe(ctx, c, id))
	assert.Contains(t, (*bodies)[1], "SUB123==")
}

func TestSyncFolderItemsPaging(t *testing.T) {
	srv, _, bodies := newServer(t, reply{200, fixture(t, "sync_page1.xml")}, reply{200, fixture(t, "sync_page2.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	ctx := context.Background()
	f := FolderRef{Id: "AAMkInbox=="}

	p1, err := SyncFolderItems(ctx, c, f, "", SyncOptions{MaxChangesReturned: 9999, IncludeParentFolderId: true, IncludeDateTimeReceived: true})
	require.NoError(t, err)
	assert.Equal(t, "STATE1", p1.SyncState)
	assert.False(t, p1.IncludesLastItemInRange)
	require.Len(t, p1.Changes, 2)
	assert.Equal(t, SyncChange{Type: SyncCreate, ItemId: "AAMkA==", ChangeKey: "CKA", ParentFolderId: "AAMkInbox==", DateTimeReceived: "2024-05-01T09:00:00Z"}, p1.Changes[0])
	assert.Equal(t, SyncChange{Type: SyncUpdate, ItemId: "AAMkB==", ChangeKey: "CKB"}, p1.Changes[1])
	assert.Contains(t, (*bodies)[0], ">512<")
	assert.Contains(t, (*bodies)[0], "item:DateTimeReceived")
	assert.NotContains(t, (*bodies)[0], "<SyncState")

	p2, err := SyncFolderItems(ctx, c, f, p1.SyncState, SyncOptions{})
	require.NoError(t, err)
	assert.True(t, p2.IncludesLastItemInRange)
	assert.Contains(t, (*bodies)[1], "STATE1")
	assert.Contains(t, (*bodies)[1], ">100<")
	require.Len(t, p2.Changes, 2)
	assert.Equal(t, SyncDelete, p2.Changes[0].Type)
	assert.Equal(t, "AAMkC==", p2.Changes[0].ItemId)
	assert.Equal(t, SyncReadFlagChange, p2.Changes[1].Type)
	require.NotNil(t, p2.Changes[1].IsRead)
	assert.True(t, *p2.Changes[1].IsRead)
}

// streamServer writes each <?xml envelope of the fixture as its own flushed
// chunk, waiting for a signal between chunks so the test proves incremental reads.
func streamServer(t *testing.T, fixtureName string, gate chan struct{}, done chan struct{}) *httptest.Server {
	envs := strings.SplitAfter(fixture(t, fixtureName), "</s:Envelope>\n")
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		defer close(done)
		fl := w.(http.Flusher)
		w.WriteHeader(200)
		for _, e := range envs {
			if strings.TrimSpace(e) == "" {
				continue
			}
			_, _ = w.Write([]byte(e))
			fl.Flush()
			if gate != nil {
				select {
				case <-gate:
				case <-r.Context().Done():
					return
				}
			}
		}
	}))
	t.Cleanup(srv.Close)
	return srv
}

func TestGetStreamingEventsIncremental(t *testing.T) {
	gate, done := make(chan struct{}), make(chan struct{})
	srv := streamServer(t, "stream_ok_notification_closed.xml", gate, done)
	c := NewContextClient(srv.URL, "u", "p", nil)

	st, err := GetStreamingEvents(context.Background(), c, []string{"SUB123=="}, 5)
	require.NoError(t, err)
	defer st.Close()

	it, err := st.Next() // first envelope read while the server is still blocked on the gate
	require.NoError(t, err)
	assert.Equal(t, ConnectionStatusOK, it.ConnectionStatus)
	assert.Equal(t, "V2017_07_11", it.ServerVersionInfo.Version)
	gate <- struct{}{}

	it, err = st.Next()
	require.NoError(t, err)
	require.Len(t, it.Notifications, 1)
	n := it.Notifications[0]
	assert.Equal(t, "SUB123==", n.SubscriptionId)
	require.Len(t, n.Events, 3)
	assert.Equal(t, StreamEventNewMail, n.Events[0].Type)
	assert.Equal(t, "2024-05-01T10:00:00Z", n.Events[0].TimeStamp)
	assert.Equal(t, &ItemId{Id: "AAMkMsg1==", ChangeKey: "CQAAAMsg1"}, n.Events[0].ItemId)
	assert.Equal(t, "AAMkInbox==", n.Events[0].ParentFolderId.Id)
	assert.Equal(t, StreamEventMoved, n.Events[1].Type)
	assert.Equal(t, "AAMkMsg2old==", n.Events[1].OldItemId.Id)
	assert.Equal(t, "AAMkInbox==", n.Events[1].OldParentFolderId.Id)
	assert.Equal(t, StreamEventModified, n.Events[2].Type)
	assert.Equal(t, "AAMkInbox==", n.Events[2].FolderId.Id)
	require.NotNil(t, n.Events[2].UnreadCount)
	assert.Equal(t, 7, *n.Events[2].UnreadCount)
	gate <- struct{}{}

	it, err = st.Next()
	require.NoError(t, err)
	assert.Equal(t, ConnectionStatusClosed, it.ConnectionStatus)
	gate <- struct{}{}

	_, err = st.Next()
	assert.Equal(t, io.EOF, err)
	<-done
}

func TestGetStreamingEventsErrorItem(t *testing.T) {
	done := make(chan struct{})
	srv := streamServer(t, "stream_error_subscription_not_found.xml", nil, done)
	c := NewContextClient(srv.URL, "u", "p", nil)
	st, err := GetStreamingEvents(context.Background(), c, []string{"SUB123==", "SUB456=="}, 1)
	require.NoError(t, err)
	defer st.Close()

	it, err := st.Next()
	require.NoError(t, err)
	require.NotNil(t, it.Err)
	assert.Equal(t, "ErrorSubscriptionNotFound", it.Err.Code)
	assert.Equal(t, []string{"SUB123==", "SUB456=="}, it.Err.SubscriptionIds)
	assert.Equal(t, "The specified subscription was not found.", it.Err.MessageText)
}

func TestGetStreamingEventsContextCancelClosesBody(t *testing.T) {
	gate, done := make(chan struct{}), make(chan struct{})
	srv := streamServer(t, "stream_ok_notification_closed.xml", gate, done)
	c := NewContextClient(srv.URL, "u", "p", nil)
	ctx, cancel := context.WithCancel(context.Background())

	st, err := GetStreamingEvents(ctx, c, []string{"SUB123=="}, 5)
	require.NoError(t, err)
	_, err = st.Next()
	require.NoError(t, err)

	cancel()
	_, err = st.Next()
	assert.True(t, errors.Is(err, context.Canceled))
	select {
	case <-done: // server saw the connection go away
	case <-time.After(3 * time.Second):
		t.Fatal("server handler still running after cancel")
	}
}

func TestGetStreamingEventsStatusErrors(t *testing.T) {
	srv, _, _ := newServer(t, reply{401, "no"})
	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := GetStreamingEvents(context.Background(), c, []string{"S"}, 5)
	assert.True(t, errors.Is(err, ErrUnauthorized))

	_, err = GetStreamingEvents(context.Background(), c, nil, 5)
	assert.Error(t, err)
	_, err = GetStreamingEvents(context.Background(), c, []string{"S"}, 31)
	assert.Error(t, err)
}

func TestHTTPClientReusedAcrossRequests(t *testing.T) {
	var conns int32
	srv := httptest.NewUnstartedServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		_, _ = w.Write([]byte(fixture(t, "unsubscribe_ok.xml")))
	}))
	srv.Config.ConnState = func(_ net.Conn, s http.ConnState) {
		if s == http.StateNew {
			atomic.AddInt32(&conns, 1)
		}
	}
	srv.Start()
	defer srv.Close()

	c := NewContextClient(srv.URL, "u", "p", nil)
	for i := 0; i < 3; i++ {
		require.NoError(t, Unsubscribe(context.Background(), c, "x"))
	}
	assert.EqualValues(t, 1, atomic.LoadInt32(&conns))
}

func TestInjectedHTTPClientAndTransport(t *testing.T) {
	srv, _, _ := newServer(t, reply{200, fixture(t, "unsubscribe_ok.xml")})

	var used int32
	hc := &http.Client{Transport: roundTripFunc(func(r *http.Request) (*http.Response, error) {
		atomic.AddInt32(&used, 1)
		return http.DefaultTransport.RoundTrip(r)
	})}
	c := NewContextClient(srv.URL, "u", "p", &Config{HTTPClient: hc})
	require.NoError(t, Unsubscribe(context.Background(), c, "x"))
	assert.EqualValues(t, 1, used)

	c = NewContextClient(srv.URL, "u", "p", &Config{Transport: hc.Transport})
	require.NoError(t, Unsubscribe(context.Background(), c, "x"))
	assert.EqualValues(t, 2, used)
}

type roundTripFunc func(*http.Request) (*http.Response, error)

func (f roundTripFunc) RoundTrip(r *http.Request) (*http.Response, error) { return f(r) }
