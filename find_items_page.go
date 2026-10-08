package ews

import (
	"context"
	"encoding/xml"
	"errors"
	"fmt"
	"time"
)

const (
	defaultFindItemsPage = 100
	// MaxFindItemsPage is the largest page size FindItemsPage requests.
	MaxFindItemsPage = 1000
)

// FindItemsOptions tunes FindItemsPage.
type FindItemsOptions struct {
	Restriction *Restriction
	// Offset and MaxEntries drive IndexedPageItemView (BasePoint=Beginning).
	// MaxEntries defaults to 100 and is capped at 1000.
	Offset     int
	MaxEntries int
	// SortBy is a FieldURI, e.g. "item:DateTimeReceived". Empty means server order.
	SortBy    string
	Ascending bool
	// AdditionalProperties are FieldURIs. item:DateTimeReceived,
	// item:ParentFolderId and message:InternetMessageId are always requested.
	AdditionalProperties []string
	// Associated searches folder-associated (FAI) items, e.g. configuration
	// items, instead of regular mail (Traversal=Associated).
	Associated bool
	// Mailbox is the SMTP address of another mailbox (delegate access). It only
	// applies to distinguished folders; pair with WithAnchorMailbox.
	Mailbox string
}

// FoundItem is one item of a FindItemsPage result.
type FoundItem struct {
	ItemId            string
	ChangeKey         string
	InternetMessageId string
	DateTimeReceived  time.Time
	ParentFolderId    string
}

// FindItemsPageResult is one page of FindItemsPage.
type FindItemsPageResult struct {
	Items                   []FoundItem
	TotalItemsInView        int
	IncludesLastItemInRange bool
	// NextOffset is the offset to request the next page (IndexedPagingOffset).
	NextOffset        int
	ServerVersionInfo ServerVersionInfo
}

type indexedPageItemView struct {
	MaxEntriesReturned int    `xml:"MaxEntriesReturned,attr"`
	Offset             int    `xml:"Offset,attr"`
	BasePoint          string `xml:"BasePoint,attr"`
}

type fieldOrder struct {
	Order    string   `xml:"Order,attr"`
	FieldURI FieldURI `xml:"http://schemas.microsoft.com/exchange/services/2006/types FieldURI"`
}

type sortOrder struct {
	FieldOrder []fieldOrder `xml:"http://schemas.microsoft.com/exchange/services/2006/types FieldOrder"`
}

// findItemsPageRequest keeps the element order the EWS schema requires.
type findItemsPageRequest struct {
	XMLName         xml.Name            `xml:"http://schemas.microsoft.com/exchange/services/2006/messages FindItem"`
	Traversal       string              `xml:"Traversal,attr"`
	ItemShape       ItemShape           `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ItemShape"`
	View            indexedPageItemView `xml:"http://schemas.microsoft.com/exchange/services/2006/messages IndexedPageItemView"`
	Restriction     *Restriction        `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Restriction,omitempty"`
	SortOrder       *sortOrder          `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SortOrder,omitempty"`
	ParentFolderIds folderRefs          `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ParentFolderIds"`
}

type foundItemXML struct {
	ItemId            ItemId  `xml:"http://schemas.microsoft.com/exchange/services/2006/types ItemId"`
	ParentFolderId    *ItemId `xml:"http://schemas.microsoft.com/exchange/services/2006/types ParentFolderId"`
	DateTimeReceived  string  `xml:"http://schemas.microsoft.com/exchange/services/2006/types DateTimeReceived"`
	InternetMessageId string  `xml:"http://schemas.microsoft.com/exchange/services/2006/types InternetMessageId"`
}

type findItemsPageEnvelope struct {
	Header serverVersionHeader `xml:"Header"`
	Body   struct {
		Response struct {
			Messages struct {
				Message struct {
					responseMessage
					RootFolder struct {
						IndexedPagingOffset     *int `xml:"IndexedPagingOffset,attr"`
						TotalItemsInView        int  `xml:"TotalItemsInView,attr"`
						IncludesLastItemInRange bool `xml:"IncludesLastItemInRange,attr"`
						Items                   struct {
							Item []foundItemXML `xml:",any"`
						} `xml:"http://schemas.microsoft.com/exchange/services/2006/types Items"`
					} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages RootFolder"`
				} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages FindItemResponseMessage"`
			} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
		} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages FindItemResponse"`
	} `xml:"Body"`
}

// FindItemsPage lists one page of items (Shallow traversal, IdOnly shape plus
// DateTimeReceived, ParentFolderId and InternetMessageId) in folder.
func FindItemsPage(ctx context.Context, c ContextClient, folder FolderRef, o FindItemsOptions, opts ...RequestOption) (*FindItemsPageResult, error) {
	if folder.isZero() {
		return nil, errors.New("ews: FindItemsPage needs a folder")
	}
	if o.Offset < 0 {
		return nil, errors.New("ews: FindItemsPage offset must not be negative")
	}
	max := o.MaxEntries
	if max <= 0 {
		max = defaultFindItemsPage
	}
	if max > MaxFindItemsPage {
		max = MaxFindItemsPage
	}

	props := &AdditionalProperties{}
	seen := map[string]bool{}
	for _, u := range append([]string{"item:DateTimeReceived", "item:ParentFolderId", "message:InternetMessageId"}, o.AdditionalProperties...) {
		if !seen[u] {
			seen[u] = true
			props.FieldURI = append(props.FieldURI, FieldURI{FieldURI: u})
		}
	}

	traversal := FindItemTraversalShallow
	if o.Associated {
		traversal = FindItemTraversalAssociated
	}
	req := findItemsPageRequest{
		Traversal:       string(traversal),
		ItemShape:       ItemShape{BaseShape: BaseShapeIdOnly, AdditionalProperties: props},
		View:            indexedPageItemView{MaxEntriesReturned: max, Offset: o.Offset, BasePoint: "Beginning"},
		Restriction:     o.Restriction,
		ParentFolderIds: newFolderRefs([]FolderRef{folder}),
	}
	if o.Mailbox != "" {
		for i := range req.ParentFolderIds.DistinguishedFolderId {
			req.ParentFolderIds.DistinguishedFolderId[i].Mailbox = &mailboxRef{EmailAddress: o.Mailbox}
		}
	}
	if o.SortBy != "" {
		order := "Descending"
		if o.Ascending {
			order = "Ascending"
		}
		req.SortOrder = &sortOrder{FieldOrder: []fieldOrder{{Order: order, FieldURI: FieldURI{FieldURI: o.SortBy}}}}
	}

	var env findItemsPageEnvelope
	if err := call(ctx, c, &req, &env, opts); err != nil {
		return nil, err
	}
	msg := env.Body.Response.Messages.Message
	if err := msg.err(); err != nil {
		return nil, err
	}

	root := msg.RootFolder
	res := &FindItemsPageResult{
		TotalItemsInView:        root.TotalItemsInView,
		IncludesLastItemInRange: root.IncludesLastItemInRange,
		ServerVersionInfo:       env.Header.ServerVersionInfo,
		Items:                   make([]FoundItem, 0, len(root.Items.Item)),
	}
	for _, it := range root.Items.Item {
		fi := FoundItem{
			ItemId:            it.ItemId.Id,
			ChangeKey:         it.ItemId.ChangeKey,
			InternetMessageId: it.InternetMessageId,
		}
		if it.ParentFolderId != nil {
			fi.ParentFolderId = it.ParentFolderId.Id
		}
		if it.DateTimeReceived != "" {
			t, err := time.Parse(time.RFC3339, it.DateTimeReceived)
			if err != nil {
				return nil, fmt.Errorf("ews: FindItemsPage item %s: bad DateTimeReceived %q: %w", it.ItemId.Id, it.DateTimeReceived, err)
			}
			fi.DateTimeReceived = t
		}
		res.Items = append(res.Items, fi)
	}
	if root.IndexedPagingOffset != nil {
		res.NextOffset = *root.IndexedPagingOffset
	} else {
		res.NextOffset = o.Offset + len(res.Items)
	}
	return res, nil
}
