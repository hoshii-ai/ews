package ews

import (
	"context"
	"errors"
	"os"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func golden(t *testing.T, name, got string) {
	t.Helper()
	if os.Getenv("UPDATE_GOLDEN") != "" {
		require.NoError(t, os.WriteFile("testdata/"+name, []byte(got+"\n"), 0o644))
	}
	assert.Equal(t, strings.TrimSpace(fixture(t, name)), got)
}

func bodyOf(t *testing.T, soap, element string) string {
	t.Helper()
	i := strings.Index(soap, "<"+element)
	j := strings.LastIndex(soap, "</"+element+">")
	require.True(t, i >= 0 && j >= 0, soap)
	return soap[i : j+len(element)+3]
}

func TestSubscribeToAllFolders(t *testing.T) {
	srv, reqs, bodies := newServer(t, reply{200, fixture(t, "subscribe_ok.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	id, err := SubscribeToAllFolders(context.Background(), c, []EventType{EventNewMail, EventCreated}, WithAnchorMailbox("a@example.com"))
	require.NoError(t, err)
	assert.NotEmpty(t, id)
	assert.Equal(t, "a@example.com", (*reqs)[0].Header.Get("X-AnchorMailbox"))

	got := bodyOf(t, (*bodies)[0], "Subscribe")
	assert.Contains(t, got, `SubscribeToAllFolders="true"`)
	assert.NotContains(t, got, "FolderIds")
	golden(t, "subscribe_all_folders_request.golden.xml", got)

	_, err = SubscribeToAllFolders(context.Background(), c, nil)
	assert.Error(t, err)
}

func TestSubscribeFolderIdsRequestUnchanged(t *testing.T) {
	srv, _, bodies := newServer(t, reply{200, fixture(t, "subscribe_ok.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := Subscribe(context.Background(), c, []FolderRef{{DistinguishedId: "inbox"}}, []EventType{EventNewMail})
	require.NoError(t, err)
	got := bodyOf(t, (*bodies)[0], "Subscribe")
	assert.NotContains(t, got, "SubscribeToAllFolders")
	assert.Contains(t, got, `DistinguishedFolderId`)
}

func TestFindFolderDeepPaged(t *testing.T) {
	srv, reqs, bodies := newServer(t, reply{200, fixture(t, "find_folder_page1.xml")}, reply{200, fixture(t, "find_folder_page2.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	res, err := FindFolder(context.Background(), c, FindFolderOptions{}, WithAnchorMailbox("a@example.com"))
	require.NoError(t, err)

	require.Len(t, res.Folders, 3)
	assert.Equal(t, FolderInfo{
		FolderId: "AAInbox", ChangeKey: "CK", ParentFolderId: "AARoot", DisplayName: "Inbox",
		FolderClass: "IPF.Note", TotalCount: 42, UnreadCount: 7, ChildFolderCount: 1,
	}, res.Folders[0])
	assert.Equal(t, "AACal", res.Folders[1].FolderId) // CalendarFolder element
	assert.Equal(t, "AAInbox", res.Folders[2].ParentFolderId)
	assert.Equal(t, "V2017_07_11", res.ServerVersionInfo.Version)

	require.Len(t, *bodies, 2)
	assert.Contains(t, (*bodies)[1], `Offset="2"`)
	assert.Equal(t, "a@example.com", (*reqs)[1].Header.Get("X-AnchorMailbox"))
	golden(t, "find_folder_request.golden.xml", bodyOf(t, (*bodies)[0], "FindFolder"))
}

func TestFindFolderServerBusy(t *testing.T) {
	srv, _, _ := newServer(t, reply{200, fixture(t, "find_folder_error_server_busy.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := FindFolder(context.Background(), c, FindFolderOptions{})
	var re *ResponseError
	require.True(t, errors.As(err, &re))
	assert.True(t, re.IsServerBusy())
}

func TestFindFolderContextCancelled(t *testing.T) {
	srv, _, _ := newServer(t, reply{200, fixture(t, "find_folder_page2.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	_, err := FindFolder(ctx, c, FindFolderOptions{})
	assert.True(t, errors.Is(err, context.Canceled))
}

func TestFindItemsPageByFolderId(t *testing.T) {
	srv, _, bodies := newServer(t, reply{200, fixture(t, "find_empty.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := FindItemsPage(context.Background(), c, FolderRef{Id: "AAFolder==", ChangeKey: "CK9"}, FindItemsOptions{})
	require.NoError(t, err)
	got := bodyOf(t, (*bodies)[0], "ParentFolderIds")
	assert.Contains(t, got, `<FolderId`)
	assert.Contains(t, got, `Id="AAFolder=="`)
	assert.Contains(t, got, `ChangeKey="CK9"`)
	assert.NotContains(t, got, "DistinguishedFolderId")
}
