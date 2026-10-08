package ews

import (
	"context"
	"encoding/xml"
	"errors"
)

// MaxSyncChanges is the largest MaxChangesReturned EWS accepts.
const MaxSyncChanges = 512

// SyncChangeType is the kind of change in a SyncFolderItems result.
type SyncChangeType string

const (
	SyncCreate         SyncChangeType = "Create"
	SyncUpdate         SyncChangeType = "Update"
	SyncDelete         SyncChangeType = "Delete"
	SyncReadFlagChange SyncChangeType = "ReadFlagChange"
)

type syncFolderItemsRequest struct {
	XMLName    xml.Name   `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SyncFolderItems"`
	ItemShape  ItemShape  `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ItemShape"`
	SyncFolder folderRefs `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SyncFolderId"`
	SyncState  string     `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SyncState,omitempty"`
	MaxChanges int        `xml:"http://schemas.microsoft.com/exchange/services/2006/messages MaxChangesReturned"`
}

type syncChangeXML struct {
	XMLName xml.Name
	// Delete and ReadFlagChange carry the ItemId directly.
	ItemId *ItemId `xml:"http://schemas.microsoft.com/exchange/services/2006/types ItemId"`
	IsRead *bool   `xml:"http://schemas.microsoft.com/exchange/services/2006/types IsRead"`
	// Create and Update wrap an item (Message, CalendarItem, ...).
	Item []struct {
		ItemId           ItemId  `xml:"http://schemas.microsoft.com/exchange/services/2006/types ItemId"`
		ParentFolderId   *ItemId `xml:"http://schemas.microsoft.com/exchange/services/2006/types ParentFolderId"`
		DateTimeReceived string  `xml:"http://schemas.microsoft.com/exchange/services/2006/types DateTimeReceived"`
	} `xml:",any"`
}

type syncFolderItemsEnvelope struct {
	Body struct {
		Response struct {
			Messages struct {
				Message struct {
					responseMessage
					SyncState               string `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SyncState"`
					IncludesLastItemInRange bool   `xml:"http://schemas.microsoft.com/exchange/services/2006/messages IncludesLastItemInRange"`
					Changes                 struct {
						Change []syncChangeXML `xml:",any"`
					} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Changes"`
				} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SyncFolderItemsResponseMessage"`
			} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
		} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SyncFolderItemsResponse"`
	} `xml:"Body"`
}

// SyncChange is one item change.
type SyncChange struct {
	Type      SyncChangeType
	ItemId    string
	ChangeKey string
	// ParentFolderId is set for Create/Update when the server returns it.
	ParentFolderId string
	// DateTimeReceived is set when requested via SyncOptions.IncludeDateTimeReceived (Create/Update).
	DateTimeReceived string
	// IsRead is set for ReadFlagChange.
	IsRead *bool
}

// SyncFolderItemsResult is one page of a SyncFolderItems sync.
type SyncFolderItemsResult struct {
	SyncState               string
	IncludesLastItemInRange bool
	Changes                 []SyncChange
}

// SyncOptions tunes SyncFolderItems.
type SyncOptions struct {
	// MaxChangesReturned caps the page size. 0 means 100; values above MaxSyncChanges are clamped.
	MaxChangesReturned int
	// IncludeParentFolderId adds item:ParentFolderId (IdOnly shape omits it).
	IncludeParentFolderId bool
	// IncludeDateTimeReceived adds item:DateTimeReceived.
	IncludeDateTimeReceived bool
}

// SyncFolderItems fetches one page of changes since syncState (empty for the
// initial sync) using the IdOnly shape. Loop until IncludesLastItemInRange,
// feeding back the returned SyncState.
func SyncFolderItems(ctx context.Context, c ContextClient, folder FolderRef, syncState string, o SyncOptions, opts ...RequestOption) (*SyncFolderItemsResult, error) {
	if folder.isZero() {
		return nil, errors.New("ews: SyncFolderItems needs a folder")
	}
	max := o.MaxChangesReturned
	if max <= 0 {
		max = 100
	}
	if max > MaxSyncChanges {
		max = MaxSyncChanges
	}
	req := syncFolderItemsRequest{
		ItemShape:  ItemShape{BaseShape: BaseShapeIdOnly},
		SyncFolder: newFolderRefs([]FolderRef{folder}),
		SyncState:  syncState,
		MaxChanges: max,
	}
	var props []FieldURI
	if o.IncludeParentFolderId {
		props = append(props, FieldURI{FieldURI: "item:ParentFolderId"})
	}
	if o.IncludeDateTimeReceived {
		props = append(props, FieldURI{FieldURI: "item:DateTimeReceived"})
	}
	if len(props) > 0 {
		req.ItemShape.AdditionalProperties = &AdditionalProperties{FieldURI: props}
	}

	var env syncFolderItemsEnvelope
	if err := call(ctx, c, &req, &env, opts); err != nil {
		return nil, err
	}
	msg := env.Body.Response.Messages.Message
	if err := msg.err(); err != nil {
		return nil, err
	}
	res := &SyncFolderItemsResult{
		SyncState:               msg.SyncState,
		IncludesLastItemInRange: msg.IncludesLastItemInRange,
	}
	for _, ch := range msg.Changes.Change {
		sc := SyncChange{Type: SyncChangeType(ch.XMLName.Local), IsRead: ch.IsRead}
		if ch.ItemId != nil {
			sc.ItemId, sc.ChangeKey = ch.ItemId.Id, ch.ItemId.ChangeKey
		}
		if len(ch.Item) > 0 {
			it := ch.Item[0]
			sc.ItemId, sc.ChangeKey = it.ItemId.Id, it.ItemId.ChangeKey
			if it.ParentFolderId != nil {
				sc.ParentFolderId = it.ParentFolderId.Id
			}
			sc.DateTimeReceived = it.DateTimeReceived
		}
		res.Changes = append(res.Changes, sc)
	}
	return res, nil
}
