package ews

import (
	"context"
	"encoding/xml"
	"errors"
)

const (
	nsMessages = "http://schemas.microsoft.com/exchange/services/2006/messages"
	nsTypes    = "http://schemas.microsoft.com/exchange/services/2006/types"
)

// call marshals req, sends it and unmarshals the response envelope into resp.
func call(ctx context.Context, c ContextClient, req, resp interface{}, opts []RequestOption) error {
	xmlBytes, err := xml.MarshalIndent(req, "", "  ")
	if err != nil {
		return err
	}
	bb, err := c.SendAndReceiveContext(ctx, xmlBytes, opts...)
	if err != nil {
		return err
	}
	return xml.Unmarshal(bb, resp)
}

// --- GetFolder ---

type getFolderRequest struct {
	XMLName     xml.Name `xml:"http://schemas.microsoft.com/exchange/services/2006/messages GetFolder"`
	FolderShape struct {
		BaseShape BaseShape `xml:"http://schemas.microsoft.com/exchange/services/2006/types BaseShape"`
	} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages FolderShape"`
	FolderIds folderRefs `xml:"http://schemas.microsoft.com/exchange/services/2006/messages FolderIds"`
}

type getFolderEnvelope struct {
	Header serverVersionHeader `xml:"Header"`
	Body   struct {
		Response struct {
			Messages struct {
				Message struct {
					responseMessage
					Folders struct {
						Folder []struct {
							FolderId folderIdRef `xml:"http://schemas.microsoft.com/exchange/services/2006/types FolderId"`
						} `xml:",any"`
					} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Folders"`
				} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages GetFolderResponseMessage"`
			} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
		} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages GetFolderResponse"`
	} `xml:"Body"`
}

// GetFolderResult is the IdOnly result of GetFolder.
type GetFolderResult struct {
	FolderId          string
	ChangeKey         string
	ServerVersionInfo ServerVersionInfo
}

// GetFolder fetches the id and change key of a folder, e.g. FolderRef{DistinguishedId: "inbox"}.
// The result also carries the server version for version detection.
func GetFolder(ctx context.Context, c ContextClient, folder FolderRef, opts ...RequestOption) (*GetFolderResult, error) {
	if folder.isZero() {
		return nil, errors.New("ews: GetFolder needs a folder")
	}
	var req getFolderRequest
	req.FolderShape.BaseShape = BaseShapeIdOnly
	req.FolderIds = newFolderRefs([]FolderRef{folder})

	var env getFolderEnvelope
	if err := call(ctx, c, &req, &env, opts); err != nil {
		return nil, err
	}
	msg := env.Body.Response.Messages.Message
	if err := msg.err(); err != nil {
		return nil, err
	}
	if len(msg.Folders.Folder) == 0 {
		return nil, errors.New("ews: GetFolder response has no folder")
	}
	f := msg.Folders.Folder[0].FolderId
	return &GetFolderResult{FolderId: f.Id, ChangeKey: f.ChangeKey, ServerVersionInfo: env.Header.ServerVersionInfo}, nil
}

// --- Subscribe (streaming) ---

// EventType is an EWS notification event type.
type EventType string

const (
	EventNewMail  EventType = "NewMailEvent"
	EventCreated  EventType = "CreatedEvent"
	EventDeleted  EventType = "DeletedEvent"
	EventModified EventType = "ModifiedEvent"
	EventMoved    EventType = "MovedEvent"
	EventCopied   EventType = "CopiedEvent"
)

type subscribeRequest struct {
	XMLName xml.Name `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Subscribe"`
	Request struct {
		SubscribeToAllFolders bool        `xml:"SubscribeToAllFolders,attr,omitempty"`
		FolderIds             *folderRefs `xml:"http://schemas.microsoft.com/exchange/services/2006/types FolderIds,omitempty"`
		EventTypes            struct {
			EventType []EventType `xml:"http://schemas.microsoft.com/exchange/services/2006/types EventType"`
		} `xml:"http://schemas.microsoft.com/exchange/services/2006/types EventTypes"`
	} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages StreamingSubscriptionRequest"`
}

type subscribeEnvelope struct {
	Body struct {
		Response struct {
			Messages struct {
				Message struct {
					responseMessage
					SubscriptionId string `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SubscriptionId"`
				} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SubscribeResponseMessage"`
			} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
		} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SubscribeResponse"`
	} `xml:"Body"`
}

// Subscribe creates a streaming subscription on the folders for the event types
// and returns the SubscriptionId.
func Subscribe(ctx context.Context, c ContextClient, folders []FolderRef, events []EventType, opts ...RequestOption) (string, error) {
	if len(folders) == 0 || len(events) == 0 {
		return "", errors.New("ews: Subscribe needs folders and event types")
	}
	refs := newFolderRefs(folders)
	var req subscribeRequest
	req.Request.FolderIds = &refs
	req.Request.EventTypes.EventType = events
	return subscribe(ctx, c, &req, opts)
}

// SubscribeToAllFolders creates a streaming subscription on every folder of the
// mailbox (Exchange 2010 SP1 and later) and returns the SubscriptionId.
func SubscribeToAllFolders(ctx context.Context, c ContextClient, events []EventType, opts ...RequestOption) (string, error) {
	if len(events) == 0 {
		return "", errors.New("ews: SubscribeToAllFolders needs event types")
	}
	var req subscribeRequest
	req.Request.SubscribeToAllFolders = true
	req.Request.EventTypes.EventType = events
	return subscribe(ctx, c, &req, opts)
}

func subscribe(ctx context.Context, c ContextClient, req *subscribeRequest, opts []RequestOption) (string, error) {
	var env subscribeEnvelope
	if err := call(ctx, c, req, &env, opts); err != nil {
		return "", err
	}
	msg := env.Body.Response.Messages.Message
	if err := msg.err(); err != nil {
		return "", err
	}
	return msg.SubscriptionId, nil
}

// --- Unsubscribe ---

type unsubscribeRequest struct {
	XMLName        xml.Name `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Unsubscribe"`
	SubscriptionId string   `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SubscriptionId"`
}

type unsubscribeEnvelope struct {
	Body struct {
		Response struct {
			Messages struct {
				Message struct {
					responseMessage
				} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages UnsubscribeResponseMessage"`
			} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
		} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages UnsubscribeResponse"`
	} `xml:"Body"`
}

// Unsubscribe ends a subscription.
func Unsubscribe(ctx context.Context, c ContextClient, subscriptionId string, opts ...RequestOption) error {
	var env unsubscribeEnvelope
	if err := call(ctx, c, &unsubscribeRequest{SubscriptionId: subscriptionId}, &env, opts); err != nil {
		return err
	}
	return env.Body.Response.Messages.Message.err()
}
