package ews

import (
	"context"
	"encoding/xml"
	"errors"
)

const findFolderPageSize = 500

// FindFolderOptions tunes FindFolder.
type FindFolderOptions struct {
	// Root is the folder to search below. Defaults to the distinguished folder msgfolderroot.
	Root FolderRef
	// Mailbox is the SMTP address of another mailbox (delegate access). It only
	// applies to a distinguished Root; pair with WithAnchorMailbox.
	Mailbox string
}

// FolderInfo describes one folder returned by FindFolder.
type FolderInfo struct {
	FolderId         string
	ChangeKey        string
	ParentFolderId   string
	DisplayName      string
	FolderClass      string
	TotalCount       int
	UnreadCount      int
	ChildFolderCount int
}

// FindFolderResult is the full folder tree below the root.
type FindFolderResult struct {
	Folders           []FolderInfo
	ServerVersionInfo ServerVersionInfo
}

type findFolderRequest struct {
	XMLName     xml.Name `xml:"http://schemas.microsoft.com/exchange/services/2006/messages FindFolder"`
	Traversal   string   `xml:"Traversal,attr"`
	FolderShape struct {
		BaseShape            BaseShape            `xml:"http://schemas.microsoft.com/exchange/services/2006/types BaseShape"`
		AdditionalProperties AdditionalProperties `xml:"http://schemas.microsoft.com/exchange/services/2006/types AdditionalProperties"`
	} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages FolderShape"`
	View            indexedPageItemView `xml:"http://schemas.microsoft.com/exchange/services/2006/messages IndexedPageFolderView"`
	ParentFolderIds folderRefs          `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ParentFolderIds"`
}

type folderXML struct {
	FolderId         folderIdRef `xml:"http://schemas.microsoft.com/exchange/services/2006/types FolderId"`
	ParentFolderId   folderIdRef `xml:"http://schemas.microsoft.com/exchange/services/2006/types ParentFolderId"`
	FolderClass      string      `xml:"http://schemas.microsoft.com/exchange/services/2006/types FolderClass"`
	DisplayName      string      `xml:"http://schemas.microsoft.com/exchange/services/2006/types DisplayName"`
	TotalCount       int         `xml:"http://schemas.microsoft.com/exchange/services/2006/types TotalCount"`
	ChildFolderCount int         `xml:"http://schemas.microsoft.com/exchange/services/2006/types ChildFolderCount"`
	UnreadCount      int         `xml:"http://schemas.microsoft.com/exchange/services/2006/types UnreadCount"`
}

type findFolderEnvelope struct {
	Header serverVersionHeader `xml:"Header"`
	Body   struct {
		Response struct {
			Messages struct {
				Message struct {
					responseMessage
					RootFolder struct {
						IndexedPagingOffset     *int `xml:"IndexedPagingOffset,attr"`
						IncludesLastItemInRange bool `xml:"IncludesLastItemInRange,attr"`
						Folders                 struct {
							Folder []folderXML `xml:",any"`
						} `xml:"http://schemas.microsoft.com/exchange/services/2006/types Folders"`
					} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages RootFolder"`
				} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages FindFolderResponseMessage"`
			} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
		} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages FindFolderResponse"`
	} `xml:"Body"`
}

// FindFolder lists every folder below o.Root (default msgfolderroot) using Deep
// traversal, paging through the server's results.
func FindFolder(ctx context.Context, c ContextClient, o FindFolderOptions, opts ...RequestOption) (*FindFolderResult, error) {
	root := o.Root
	if root.isZero() {
		root = FolderRef{DistinguishedId: "msgfolderroot"}
	}
	refs := newFolderRefs([]FolderRef{root})
	if o.Mailbox != "" {
		for i := range refs.DistinguishedFolderId {
			refs.DistinguishedFolderId[i].Mailbox = &mailboxRef{EmailAddress: o.Mailbox}
		}
	}

	res := &FindFolderResult{}
	offset := 0
	for {
		var req findFolderRequest
		req.Traversal = "Deep"
		req.FolderShape.BaseShape = BaseShapeIdOnly
		for _, u := range []string{
			"folder:ParentFolderId", "folder:DisplayName", "folder:FolderClass",
			"folder:TotalCount", "folder:UnreadCount", "folder:ChildFolderCount",
		} {
			req.FolderShape.AdditionalProperties.FieldURI = append(req.FolderShape.AdditionalProperties.FieldURI, FieldURI{FieldURI: u})
		}
		req.View = indexedPageItemView{MaxEntriesReturned: findFolderPageSize, Offset: offset, BasePoint: "Beginning"}
		req.ParentFolderIds = refs

		var env findFolderEnvelope
		if err := call(ctx, c, &req, &env, opts); err != nil {
			return nil, err
		}
		msg := env.Body.Response.Messages.Message
		if err := msg.err(); err != nil {
			return nil, err
		}
		res.ServerVersionInfo = env.Header.ServerVersionInfo

		page := msg.RootFolder.Folders.Folder
		for _, f := range page {
			res.Folders = append(res.Folders, FolderInfo{
				FolderId:         f.FolderId.Id,
				ChangeKey:        f.FolderId.ChangeKey,
				ParentFolderId:   f.ParentFolderId.Id,
				DisplayName:      f.DisplayName,
				FolderClass:      f.FolderClass,
				TotalCount:       f.TotalCount,
				UnreadCount:      f.UnreadCount,
				ChildFolderCount: f.ChildFolderCount,
			})
		}
		if msg.RootFolder.IncludesLastItemInRange || len(page) == 0 {
			return res, nil
		}
		next := offset + len(page)
		if p := msg.RootFolder.IndexedPagingOffset; p != nil {
			next = *p
		}
		if next <= offset {
			return nil, errors.New("ews: FindFolder paging made no progress")
		}
		offset = next
	}
}
