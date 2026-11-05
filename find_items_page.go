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
	// MaxFindItemsPage is the largest page size FindItemsPage and FindItemsPages request.
	MaxFindItemsPage = 1000
)

// FindItemsOptions tunes FindItemsPage and FindItemsPages.
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
	Items []FoundItem
	// TotalItemsInView is not a count once the offset is past the folder's
	// last item: Exchange then echoes the offset.
	TotalItemsInView int
	// IncludesLastItemInRange can stay false when a page ends exactly on the
	// folder's last item. Use Done to stop paging.
	IncludesLastItemInRange bool
	// Done reports that the page came back shorter than the MaxEntries
	// requested and IncludesLastItemInRange is true, so the folder has no item
	// past it. A short page without IncludesLastItemInRange was cut by the
	// server (for example under a find count throttling limit) and is not Done.
	Done bool
	// NextOffset is the offset to request the next page (IndexedPagingOffset).
	NextOffset        int
	ServerVersionInfo ServerVersionInfo
}

// FolderItemsPage is one folder's result of FindItemsPages: its Page, or Err
// when the folder's own response message is an error (e.g. ErrorFolderNotFound).
type FolderItemsPage struct {
	Folder FolderRef
	Page   *FindItemsPageResult
	Err    error
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

type findItemsPageMessage struct {
	responseMessage
	RootFolder struct {
		IndexedPagingOffset     *int `xml:"IndexedPagingOffset,attr"`
		TotalItemsInView        int  `xml:"TotalItemsInView,attr"`
		IncludesLastItemInRange bool `xml:"IncludesLastItemInRange,attr"`
		Items                   struct {
			Item []foundItemXML `xml:",any"`
		} `xml:"http://schemas.microsoft.com/exchange/services/2006/types Items"`
	} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages RootFolder"`
}

// findItemsPageEnvelope holds one response message per parent folder, in request order.
type findItemsPageEnvelope struct {
	Header serverVersionHeader `xml:"Header"`
	Body   struct {
		Response struct {
			Messages struct {
				Message []findItemsPageMessage `xml:"http://schemas.microsoft.com/exchange/services/2006/messages FindItemResponseMessage"`
			} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
		} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages FindItemResponse"`
	} `xml:"Body"`
}

// FindItemsPage lists one page of items (Shallow traversal, IdOnly shape plus
// DateTimeReceived, ParentFolderId and InternetMessageId) in folder.
func FindItemsPage(ctx context.Context, c ContextClient, folder FolderRef, o FindItemsOptions, opts ...RequestOption) (*FindItemsPageResult, error) {
	pages, err := FindItemsPages(ctx, c, []FolderRef{folder}, o, opts...)
	if err != nil {
		return nil, err
	}
	return pages[0].Page, nil
}

// FindItemsPages lists one page of items in each folder with one FindItem
// request. Exchange answers one response message per parent folder, each sorted
// and paged on its own and never merged: the restriction, the sort and the
// offset apply to every folder, so the request returns page k of each folder.
// The result follows the order of folders. A folder whose response message is
// an error carries it in Err and the others keep their page; the call fails
// only when the request fails as a whole or every folder failed.
func FindItemsPages(ctx context.Context, c ContextClient, folders []FolderRef, o FindItemsOptions, opts ...RequestOption) ([]FolderItemsPage, error) {
	if len(folders) == 0 {
		return nil, errors.New("ews: FindItemsPages needs a folder")
	}
	for _, f := range folders {
		if f.isZero() {
			return nil, errors.New("ews: FindItemsPages needs non-empty folders")
		}
	}
	if o.Offset < 0 {
		return nil, errors.New("ews: FindItemsPages offset must not be negative")
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
		ParentFolderIds: newFolderRefs(folders),
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
	msgs := env.Body.Response.Messages.Message
	if err := checkFolderResponses(msgs, len(folders)); err != nil {
		return nil, err
	}

	// The request lists distinguished folders before folder ids, and the response messages follow it.
	nextDistinguished, nextId := 0, len(req.ParentFolderIds.DistinguishedFolderId)
	pages := make([]FolderItemsPage, len(folders))
	var firstErr error
	failed := 0
	for i, f := range folders {
		var msg findItemsPageMessage
		if f.DistinguishedId != "" {
			msg, nextDistinguished = msgs[nextDistinguished], nextDistinguished+1
		} else {
			msg, nextId = msgs[nextId], nextId+1
		}
		pages[i].Folder = f
		if err := msg.err(); err != nil {
			pages[i].Err = err
			if firstErr == nil {
				firstErr = err
			}
			failed++
			continue
		}
		page, err := msg.page(o.Offset, max, env.Header.ServerVersionInfo)
		if err != nil {
			return nil, err
		}
		pages[i].Page = page
	}
	if failed == len(folders) {
		return nil, firstErr
	}
	return pages, nil
}

// checkFolderResponses fails on a request-level error, which is one error
// message for several folders (ErrorServerBusy, a server refusing several
// folders), and on a response count that does not match the n folders asked.
func checkFolderResponses(msgs []findItemsPageMessage, n int) error {
	if len(msgs) == 1 && n > 1 {
		if err := msgs[0].err(); err != nil {
			return err
		}
	}
	if len(msgs) != n {
		return fmt.Errorf("ews: FindItem asked %d folders, got %d response messages", n, len(msgs))
	}
	return nil
}

func (m findItemsPageMessage) page(offset, max int, version ServerVersionInfo) (*FindItemsPageResult, error) {
	items, err := m.foundItems()
	if err != nil {
		return nil, err
	}
	root := m.RootFolder
	res := &FindItemsPageResult{
		TotalItemsInView:        root.TotalItemsInView,
		IncludesLastItemInRange: root.IncludesLastItemInRange,
		Done:                    len(items) < max && root.IncludesLastItemInRange,
		ServerVersionInfo:       version,
		Items:                   items,
	}
	if root.IndexedPagingOffset != nil {
		res.NextOffset = *root.IndexedPagingOffset
	} else {
		res.NextOffset = offset + len(items)
	}
	return res, nil
}

func (m findItemsPageMessage) foundItems() ([]FoundItem, error) {
	items := make([]FoundItem, 0, len(m.RootFolder.Items.Item))
	for _, it := range m.RootFolder.Items.Item {
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
		items = append(items, fi)
	}
	return items, nil
}
