package ews

import (
	"context"
	"encoding/base64"
	"errors"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestMoveItem(t *testing.T) {
	srv, reqs, bodies := newServer(t, reply{200, fixture(t, "move_item_ok.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	got, err := MoveItem(context.Background(), c, ItemId{Id: "AAMk1", ChangeKey: "CK1"}, FolderRef{Id: "AAFolder==", ChangeKey: "CKF"}, WithAnchorMailbox("a@example.com"))
	require.NoError(t, err)
	assert.Equal(t, &ItemId{Id: "AAMkNew==", ChangeKey: "CKNew"}, got)
	assert.Equal(t, "a@example.com", (*reqs)[0].Header.Get("X-AnchorMailbox"))
	golden(t, "move_item_request.golden.xml", bodyOf(t, (*bodies)[0], "MoveItem"))

	// distinguished destination
	_, err = MoveItem(context.Background(), c, ItemId{Id: "AAMk1"}, FolderRef{DistinguishedId: "deleteditems"})
	require.NoError(t, err)
	assert.Contains(t, (*bodies)[1], `<DistinguishedFolderId xmlns="http://schemas.microsoft.com/exchange/services/2006/types" Id="deleteditems">`)

	_, err = MoveItem(context.Background(), c, ItemId{}, FolderRef{Id: "x"})
	assert.Error(t, err)
	_, err = MoveItem(context.Background(), c, ItemId{Id: "x"}, FolderRef{})
	assert.Error(t, err)
}

func TestMoveItemServerBusy(t *testing.T) {
	srv, _, _ := newServer(t, reply{200, fixture(t, "move_item_busy.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := MoveItem(context.Background(), c, ItemId{Id: "a"}, FolderRef{Id: "b"})
	var re *ResponseError
	require.True(t, errors.As(err, &re))
	assert.True(t, re.IsServerBusy())
	assert.Equal(t, 1000, re.BackOffMilliseconds)
}

func TestFindItemByInternetMessageID(t *testing.T) {
	srv, _, bodies := newServer(t, reply{200, fixture(t, "find_msgid_match.xml")}, reply{200, fixture(t, "find_empty.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	items, err := FindItemByInternetMessageID(context.Background(), c, FolderRef{DistinguishedId: "inbox"}, "<a@x>")
	require.NoError(t, err)
	require.Len(t, items, 2)
	assert.Equal(t, "AAMk1", items[0].ItemId)
	assert.Equal(t, "AAMkInbox==", items[0].ParentFolderId)
	golden(t, "find_by_message_id_request.golden.xml", bodyOf(t, (*bodies)[0], "FindItem"))

	// zero matches is not an error
	items, err = FindItemByInternetMessageID(context.Background(), c, FolderRef{Id: "AAFolder=="}, "<none@x>")
	require.NoError(t, err)
	assert.Empty(t, items)

	_, err = FindItemByInternetMessageID(context.Background(), c, FolderRef{DistinguishedId: "inbox"}, "")
	assert.Error(t, err)
}

func TestSendMIME(t *testing.T) {
	srv, reqs, bodies := newServer(t, reply{200, fixture(t, "create_item_mime_ok.xml")}, reply{200, fixture(t, "create_item_mime_noid.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	mime := []byte("Message-ID: <r1@example.com>\r\nIn-Reply-To: <a@x>\r\nReferences: <a@x>\r\nSubject: Re: hi\r\n\r\nthanks\r\n")
	id, err := SendMIME(context.Background(), c, mime, FolderRef{DistinguishedId: "sentitems"}, WithAnchorMailbox("a@example.com"))
	require.NoError(t, err)
	assert.Equal(t, &ItemId{Id: "AAMkSent==", ChangeKey: "CKS"}, id)
	assert.Equal(t, "a@example.com", (*reqs)[0].Header.Get("X-AnchorMailbox"))
	golden(t, "send_mime_request.golden.xml", bodyOf(t, (*bodies)[0], "CreateItem"))
	assert.Contains(t, (*bodies)[0], base64.StdEncoding.EncodeToString(mime))
	assert.Contains(t, (*bodies)[0], `MessageDisposition="SendAndSaveCopy"`)

	// no saveTo and no id in the response
	id, err = SendMIME(context.Background(), c, mime, FolderRef{})
	require.NoError(t, err)
	assert.Nil(t, id)
	assert.NotContains(t, (*bodies)[1], "SavedItemFolderId")

	_, err = SendMIME(context.Background(), c, nil, FolderRef{})
	assert.Error(t, err)
}

func TestGetItemContext(t *testing.T) {
	srv, reqs, _ := newServer(t, reply{200, fixture(t, "get_item_mime.xml")}, reply{200, fixture(t, "get_item_error.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	resp, err := GetItemContext(context.Background(), c, ItemId{Id: "AAMk1"}, GetItemRequestConfig{ItemShape: NewGetItemMimeShape()}, WithAnchorMailbox("a@example.com"))
	require.NoError(t, err)
	assert.NotNil(t, resp.ResponseMessages.GetItemResponseMessage.Items.Message[0].MimeContent)
	assert.Equal(t, "a@example.com", (*reqs)[0].Header.Get("X-AnchorMailbox"))

	_, err = GetItemContext(context.Background(), c, ItemId{Id: "AAMk1"}, GetItemRequestConfig{})
	var re *ResponseError
	require.True(t, errors.As(err, &re))
	assert.Equal(t, "ErrorAccessDenied", re.Code)
	assert.Equal(t, "Access is denied.", re.MessageText)
}

func TestUpdateItemContext(t *testing.T) {
	srv, reqs, bodies := newServer(t, reply{200, fixture(t, "update_item_ok.xml")}, reply{200, fixture(t, "update_item_error.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	subject := "new"
	r := &UpdateItemRequest{
		MessageDisposition: MessageDispositionSaveOnly,
		ItemChanges: ItemChanges{ItemChange: []ItemChange{{
			ItemId: ItemId{Id: "AAMk1", ChangeKey: "CK1"},
			Updates: Updates{SetItemField: []SetItemField{{
				FieldURI: &FieldURI{FieldURI: "item:Subject"},
				Message:  &Message{Subject: &subject},
			}}},
		}}},
	}
	_, err := UpdateItemContext(context.Background(), c, r, WithAnchorMailbox("a@example.com"))
	require.NoError(t, err)
	assert.Equal(t, "a@example.com", (*reqs)[0].Header.Get("X-AnchorMailbox"))
	golden(t, "update_item_request.golden.xml", bodyOf(t, (*bodies)[0], "UpdateItem"))

	_, err = UpdateItemContext(context.Background(), c, r)
	var re *ResponseError
	require.True(t, errors.As(err, &re))
	assert.Equal(t, "ErrorAccessDenied", re.Code)
}

func TestSendItemChecksResponseClass(t *testing.T) {
	srv, _, bodies := newServer(t, reply{200, fixture(t, "send_item_ok.xml")}, reply{200, fixture(t, "send_item_error.xml")}, reply{200, fixture(t, "send_item_error.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	ctx := context.Background()

	require.NoError(t, SendItemContext(ctx, c, ItemId{Id: "AAMk1", ChangeKey: "CK1"}, true))
	golden(t, "send_item_request.golden.xml", bodyOf(t, (*bodies)[0], "SendItem"))

	var re *ResponseError
	err := SendItemContext(ctx, c, ItemId{Id: "AAMk1"}, true)
	require.True(t, errors.As(err, &re))
	assert.Equal(t, "ErrorItemNotFound", re.Code)

	// legacy SendItem used to swallow Error responses
	_, err = SendItem(c, ItemId{Id: "AAMk1"}, true)
	require.True(t, errors.As(err, &re))
	assert.Equal(t, "ErrorItemNotFound", re.Code)
}

func TestFindItemsPageAssociated(t *testing.T) {
	srv, _, bodies := newServer(t, reply{200, fixture(t, "find_empty.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := FindItemsPage(context.Background(), c, FolderRef{DistinguishedId: "calendar"}, FindItemsOptions{Associated: true})
	require.NoError(t, err)
	assert.Contains(t, (*bodies)[0], `Traversal="Associated"`)
}
