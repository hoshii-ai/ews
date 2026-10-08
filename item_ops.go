package ews

import (
	"context"
	"encoding/base64"
	"encoding/xml"
	"errors"
)

// --- MoveItem ---

type moveItemRequest struct {
	XMLName    xml.Name   `xml:"http://schemas.microsoft.com/exchange/services/2006/messages MoveItem"`
	ToFolderId folderRefs `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ToFolderId"`
	ItemIds    ItemIds    `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ItemIds"`
}

type moveItemEnvelope struct {
	Body struct {
		Response struct {
			Messages struct {
				Message struct {
					responseMessage
					Items struct {
						Item []struct {
							ItemId ItemId `xml:"http://schemas.microsoft.com/exchange/services/2006/types ItemId"`
						} `xml:",any"`
					} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Items"`
				} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages MoveItemResponseMessage"`
			} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
		} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages MoveItemResponse"`
	} `xml:"Body"`
}

// MoveItem moves an item to another folder and returns the item's new id and
// change key. to is a FolderRef with either a DistinguishedId or an Id.
func MoveItem(ctx context.Context, c ContextClient, item ItemId, to FolderRef, opts ...RequestOption) (*ItemId, error) {
	if item.Id == "" {
		return nil, errors.New("ews: MoveItem needs an item id")
	}
	if to.isZero() {
		return nil, errors.New("ews: MoveItem needs a destination folder")
	}
	req := moveItemRequest{
		ToFolderId: newFolderRefs([]FolderRef{to}),
		ItemIds:    ItemIds{ItemId: []ItemId{item}},
	}
	var env moveItemEnvelope
	if err := call(ctx, c, &req, &env, opts); err != nil {
		return nil, err
	}
	msg := env.Body.Response.Messages.Message
	if err := msg.err(); err != nil {
		return nil, err
	}
	if len(msg.Items.Item) == 0 || msg.Items.Item[0].ItemId.Id == "" {
		return nil, errors.New("ews: MoveItem response has no item id")
	}
	moved := msg.Items.Item[0].ItemId
	return &moved, nil
}

// --- FindItemByInternetMessageID ---

// FindItemByInternetMessageID looks up items in folder whose Internet
// Message-ID equals id (as stored, normally including the angle brackets).
// No match is not an error: the result is empty.
func FindItemByInternetMessageID(ctx context.Context, c ContextClient, folder FolderRef, id string, opts ...RequestOption) ([]FoundItem, error) {
	if id == "" {
		return nil, errors.New("ews: FindItemByInternetMessageID needs a message id")
	}
	restriction := &Restriction{IsEqualTo: &IsEqualTo{
		ExtendedFieldURI:   &ExtendedFieldURI{PropertyTag: PropertyTagInternetMessageId, PropertyType: PropertyTypeString},
		FieldURIOrConstant: &FieldURIOrConstant{Constant: &Constant{Value: id}},
	}}
	res, err := FindItemsPage(ctx, c, folder, FindItemsOptions{Restriction: restriction}, opts...)
	if err != nil {
		return nil, err
	}
	return res.Items, nil
}

// --- SendMIME ---

type sendMIMERequest struct {
	XMLName            xml.Name           `xml:"http://schemas.microsoft.com/exchange/services/2006/messages CreateItem"`
	MessageDisposition MessageDisposition `xml:"MessageDisposition,attr"`
	SavedItemFolderId  *folderRefs        `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SavedItemFolderId,omitempty"`
	Items              struct {
		Message struct {
			MimeContent MimeContent `xml:"http://schemas.microsoft.com/exchange/services/2006/types MimeContent"`
		} `xml:"http://schemas.microsoft.com/exchange/services/2006/types Message"`
	} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Items"`
}

type sendMIMEEnvelope struct {
	Body struct {
		Response struct {
			Messages struct {
				Message struct {
					responseMessage
					Items struct {
						Item []struct {
							ItemId *ItemId `xml:"http://schemas.microsoft.com/exchange/services/2006/types ItemId"`
						} `xml:",any"`
					} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Items"`
				} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages CreateItemResponseMessage"`
			} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
		} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages CreateItemResponse"`
	} `xml:"Body"`
}

// SendMIME sends a complete RFC 5322 message (headers, body and attachments)
// as is, so headers such as Message-ID, In-Reply-To and References are
// carried in the MIME stream, and saves a copy in saveTo (SendAndSaveCopy).
// A zero saveTo lets the server pick the Sent Items folder. The returned
// ItemId is the saved copy, or nil when the server does not report one.
func SendMIME(ctx context.Context, c ContextClient, mime []byte, saveTo FolderRef, opts ...RequestOption) (*ItemId, error) {
	if len(mime) == 0 {
		return nil, errors.New("ews: SendMIME needs a MIME message")
	}
	req := sendMIMERequest{MessageDisposition: MessageDispositionSendAndSaveCopy}
	req.Items.Message.MimeContent = MimeContent{CharacterSet: "UTF-8", Value: base64.StdEncoding.EncodeToString(mime)}
	if !saveTo.isZero() {
		refs := newFolderRefs([]FolderRef{saveTo})
		req.SavedItemFolderId = &refs
	}
	var env sendMIMEEnvelope
	if err := call(ctx, c, &req, &env, opts); err != nil {
		return nil, err
	}
	msg := env.Body.Response.Messages.Message
	if err := msg.err(); err != nil {
		return nil, err
	}
	if len(msg.Items.Item) > 0 {
		return msg.Items.Item[0].ItemId, nil
	}
	return nil, nil
}
