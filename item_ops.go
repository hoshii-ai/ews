package ews

import (
	"context"
	"encoding/base64"
	"encoding/xml"
	"errors"
	"fmt"
)

// --- MoveItem ---

type moveItemRequest struct {
	XMLName xml.Name `xml:"http://schemas.microsoft.com/exchange/services/2006/messages MoveItem"`
	// ReturnNewItemIds asks for the moved item's new id; without it Exchange 2013+
	// answers Success with an empty Items element.
	ReturnNewItemIds bool       `xml:"ReturnNewItemIds,attr"`
	ToFolderId       folderRefs `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ToFolderId"`
	ItemIds          ItemIds    `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ItemIds"`
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
// A successful move is never an error: when the server reports no new id
// (older servers, some cross-mailbox moves) the result is (nil, nil).
func MoveItem(ctx context.Context, c ContextClient, item ItemId, to FolderRef, opts ...RequestOption) (*ItemId, error) {
	if item.Id == "" {
		return nil, errors.New("ews: MoveItem needs an item id")
	}
	if to.isZero() {
		return nil, errors.New("ews: MoveItem needs a destination folder")
	}
	req := moveItemRequest{
		ReturnNewItemIds: true,
		ToFolderId:       newFolderRefs([]FolderRef{to}),
		ItemIds:          ItemIds{ItemId: []ItemId{item}},
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
		return nil, nil
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
	res, err := FindItemsPage(ctx, c, folder, FindItemsOptions{Restriction: internetMessageIDRestriction(id)}, opts...)
	if err != nil {
		return nil, err
	}
	return res.Items, nil
}

// --- FindItemsByInternetMessageID ---

// FolderItems is one folder's result of FindItemsByInternetMessageID. Err is
// that folder's own response error (e.g. ErrorFolderNotFound); the other
// folders are still answered.
type FolderItems struct {
	Folder FolderRef
	Items  []FoundItem
	Err    error
}

// FindItemsByInternetMessageID looks up id in every folder with one FindItem
// request. FindItem has no Deep traversal, but it takes several parent folders
// and answers one response message per folder. The result follows the
// serialized request order, which lists distinguished folders before folder ids.
func FindItemsByInternetMessageID(ctx context.Context, c ContextClient, folders []FolderRef, id string, opts ...RequestOption) ([]FolderItems, error) {
	if id == "" {
		return nil, errors.New("ews: FindItemsByInternetMessageID needs a message id")
	}
	for _, f := range folders {
		if f.isZero() {
			return nil, errors.New("ews: FindItemsByInternetMessageID needs non-empty folders")
		}
	}
	if len(folders) == 0 {
		return nil, errors.New("ews: FindItemsByInternetMessageID needs a folder")
	}
	refs := newFolderRefs(folders)
	// The responses follow the serialized order, so read it back from refs rather than assume it.
	ordered := make([]FolderRef, 0, len(folders))
	for _, d := range refs.DistinguishedFolderId {
		ordered = append(ordered, FolderRef{DistinguishedId: d.Id})
	}
	for _, f := range refs.FolderId {
		ordered = append(ordered, FolderRef{Id: f.Id, ChangeKey: f.ChangeKey})
	}

	req := findItemsPageRequest{
		Traversal: string(FindItemTraversalShallow),
		ItemShape: ItemShape{BaseShape: BaseShapeIdOnly, AdditionalProperties: &AdditionalProperties{FieldURI: []FieldURI{
			{FieldURI: "item:DateTimeReceived"},
			{FieldURI: "item:ParentFolderId"},
			{FieldURI: "message:InternetMessageId"},
		}}},
		View:            indexedPageItemView{MaxEntriesReturned: defaultFindItemsPage, BasePoint: "Beginning"},
		Restriction:     internetMessageIDRestriction(id),
		ParentFolderIds: refs,
	}
	var env findItemsPageEnvelope
	if err := call(ctx, c, &req, &env, opts); err != nil {
		return nil, err
	}
	msgs := env.Body.Response.Messages.Message
	// A request-level failure (ErrorServerBusy, a server refusing several folders) is one error message.
	if len(msgs) == 1 && len(ordered) > 1 {
		if err := msgs[0].err(); err != nil {
			return nil, err
		}
	}
	if len(msgs) != len(ordered) {
		return nil, fmt.Errorf("ews: FindItemsByInternetMessageID asked %d folders, got %d response messages", len(ordered), len(msgs))
	}

	res := make([]FolderItems, len(ordered))
	for i, msg := range msgs {
		res[i].Folder = ordered[i]
		if err := msg.err(); err != nil {
			res[i].Err = err
			continue
		}
		items, err := msg.foundItems()
		if err != nil {
			return nil, err
		}
		res[i].Items = items
	}
	return res, nil
}

func internetMessageIDRestriction(id string) *Restriction {
	return &Restriction{IsEqualTo: &IsEqualTo{
		ExtendedFieldURI:   &ExtendedFieldURI{PropertyTag: PropertyTagInternetMessageId, PropertyType: PropertyTypeString},
		FieldURIOrConstant: &FieldURIOrConstant{Constant: &Constant{Value: id}},
	}}
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
