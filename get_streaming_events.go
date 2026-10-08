package ews

import (
	"context"
	"encoding/xml"
	"errors"
	"io"
	"strings"
)

// StreamEventType is the kind of an event inside a notification.
type StreamEventType string

const (
	StreamEventNewMail  StreamEventType = "NewMailEvent"
	StreamEventCreated  StreamEventType = "CreatedEvent"
	StreamEventDeleted  StreamEventType = "DeletedEvent"
	StreamEventModified StreamEventType = "ModifiedEvent"
	StreamEventMoved    StreamEventType = "MovedEvent"
	StreamEventCopied   StreamEventType = "CopiedEvent"
)

// Connection status values reported by the stream.
const (
	ConnectionStatusOK     = "OK"
	ConnectionStatusClosed = "Closed"
)

type getStreamingEventsRequest struct {
	XMLName         xml.Name `xml:"http://schemas.microsoft.com/exchange/services/2006/messages GetStreamingEvents"`
	SubscriptionIds struct {
		Id []string `xml:"http://schemas.microsoft.com/exchange/services/2006/types SubscriptionId"`
	} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SubscriptionIds"`
	ConnectionTimeout int `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ConnectionTimeout"`
}

type streamEventXML struct {
	XMLName           xml.Name
	Watermark         string  `xml:"http://schemas.microsoft.com/exchange/services/2006/types Watermark"`
	TimeStamp         string  `xml:"http://schemas.microsoft.com/exchange/services/2006/types TimeStamp"`
	ItemId            *ItemId `xml:"http://schemas.microsoft.com/exchange/services/2006/types ItemId"`
	FolderId          *ItemId `xml:"http://schemas.microsoft.com/exchange/services/2006/types FolderId"`
	ParentFolderId    *ItemId `xml:"http://schemas.microsoft.com/exchange/services/2006/types ParentFolderId"`
	OldItemId         *ItemId `xml:"http://schemas.microsoft.com/exchange/services/2006/types OldItemId"`
	OldFolderId       *ItemId `xml:"http://schemas.microsoft.com/exchange/services/2006/types OldFolderId"`
	OldParentFolderId *ItemId `xml:"http://schemas.microsoft.com/exchange/services/2006/types OldParentFolderId"`
	UnreadCount       *int    `xml:"http://schemas.microsoft.com/exchange/services/2006/types UnreadCount"`
}

type streamNotificationXML struct {
	SubscriptionId string           `xml:"http://schemas.microsoft.com/exchange/services/2006/types SubscriptionId"`
	Events         []streamEventXML `xml:",any"`
}

type streamMessageXML struct {
	responseMessage
	ConnectionStatus string `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ConnectionStatus"`
	Notifications    struct {
		Notification []streamNotificationXML `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Notification"`
	} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Notifications"`
}

type streamEnvelopeXML struct {
	Header serverVersionHeader `xml:"Header"`
	Body   struct {
		Response struct {
			Messages struct {
				Message []streamMessageXML `xml:"http://schemas.microsoft.com/exchange/services/2006/messages GetStreamingEventsResponseMessage"`
			} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
		} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages GetStreamingEventsResponse"`
	} `xml:"Body"`
}

// StreamEvent is one event of a notification. Pointer-free fields are empty when absent.
type StreamEvent struct {
	Type      StreamEventType
	Watermark string
	TimeStamp string // as sent, e.g. 2024-05-01T10:00:00Z
	// ItemId is set for item events; FolderId for folder events.
	ItemId   *ItemId
	FolderId *ItemId
	// ParentFolderId is the folder the item/folder lives in (the new one for Moved/Copied).
	ParentFolderId *ItemId
	// Old* are set for Moved/Copied events.
	OldItemId         *ItemId
	OldFolderId       *ItemId
	OldParentFolderId *ItemId
	// UnreadCount is set on events that report it (nil otherwise).
	UnreadCount *int
}

// Notification groups the events of one subscription.
type Notification struct {
	SubscriptionId string
	Events         []StreamEvent
}

// StreamItem is one response message from the stream. Exactly one of the
// following is meaningful: Notifications non-empty, ConnectionStatus non-empty
// (OK keepalive, Closed), or Err non-nil.
type StreamItem struct {
	Notifications    []Notification
	ConnectionStatus string
	// Err is a *ResponseError (Code, MessageText, BackOffMilliseconds, SubscriptionIds).
	Err               *ResponseError
	ServerVersionInfo ServerVersionInfo
}

// EventStream reads the open GetStreamingEvents response body envelope by envelope.
type EventStream struct {
	body    io.ReadCloser
	dec     *xml.Decoder
	ctx     context.Context
	pending []StreamItem
}

// GetStreamingEvents opens a streaming connection for the subscriptions. The
// server keeps it open for connectionTimeoutMinutes (1..30) and sends keepalives.
// Close the returned stream (or cancel ctx) to release the connection.
// The client's http.Client must have Timeout == 0 (a Timeout would kill the
// stream); bound the stream with ctx instead.
func GetStreamingEvents(ctx context.Context, c ContextClient, subscriptionIds []string, connectionTimeoutMinutes int, opts ...RequestOption) (*EventStream, error) {
	if len(subscriptionIds) == 0 {
		return nil, errors.New("ews: GetStreamingEvents needs subscription ids")
	}
	if connectionTimeoutMinutes < 1 || connectionTimeoutMinutes > 30 {
		return nil, errors.New("ews: connection timeout must be 1..30 minutes")
	}
	var req getStreamingEventsRequest
	req.SubscriptionIds.Id = subscriptionIds
	req.ConnectionTimeout = connectionTimeoutMinutes
	xmlBytes, err := xml.MarshalIndent(&req, "", "  ")
	if err != nil {
		return nil, err
	}
	body, err := c.OpenStream(ctx, xmlBytes, opts...)
	if err != nil {
		return nil, err
	}
	return newEventStream(ctx, body), nil
}

func newEventStream(ctx context.Context, body io.ReadCloser) *EventStream {
	return &EventStream{body: body, dec: xml.NewDecoder(body), ctx: ctx}
}

// Next blocks until the next item. It returns io.EOF when the server ends the
// stream (normally after a Closed status), ctx.Err() after cancellation, or a
// parse/transport error. EWS-level errors come back as StreamItem.Err, not as error.
func (s *EventStream) Next() (StreamItem, error) {
	for len(s.pending) == 0 {
		var env streamEnvelopeXML
		if err := s.dec.Decode(&env); err != nil {
			if cerr := s.ctx.Err(); cerr != nil {
				return StreamItem{}, cerr
			}
			return StreamItem{}, err
		}
		s.pending = append(s.pending, itemsFromEnvelope(&env)...)
	}
	it := s.pending[0]
	s.pending = s.pending[1:]
	return it, nil
}

// Close closes the underlying response body.
func (s *EventStream) Close() error {
	return s.body.Close()
}

func itemsFromEnvelope(env *streamEnvelopeXML) []StreamItem {
	var items []StreamItem
	for _, m := range env.Body.Response.Messages.Message {
		it := StreamItem{ServerVersionInfo: env.Header.ServerVersionInfo}
		if err := m.err(); err != nil {
			it.Err = err.(*ResponseError)
			items = append(items, it)
			continue
		}
		it.ConnectionStatus = m.ConnectionStatus
		for _, n := range m.Notifications.Notification {
			notif := Notification{SubscriptionId: n.SubscriptionId}
			for _, e := range n.Events {
				if !strings.HasSuffix(e.XMLName.Local, "Event") {
					continue // PreviousWatermark, MoreEvents, ...
				}
				notif.Events = append(notif.Events, StreamEvent{
					Type:              StreamEventType(e.XMLName.Local),
					Watermark:         e.Watermark,
					TimeStamp:         e.TimeStamp,
					ItemId:            e.ItemId,
					FolderId:          e.FolderId,
					ParentFolderId:    e.ParentFolderId,
					OldItemId:         e.OldItemId,
					OldFolderId:       e.OldFolderId,
					OldParentFolderId: e.OldParentFolderId,
					UnreadCount:       e.UnreadCount,
				})
			}
			it.Notifications = append(it.Notifications, notif)
		}
		items = append(items, it)
	}
	return items
}
