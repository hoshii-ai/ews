package ews

import (
	"context"
	"encoding/xml"
	"errors"
	"os"
	"regexp"
	"strings"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// requestBody extracts the FindItem element sent inside the SOAP envelope.
func requestBody(t *testing.T, soap string) string {
	t.Helper()
	i := strings.Index(soap, "<FindItem")
	require.GreaterOrEqual(t, i, 0, soap)
	j := strings.LastIndex(soap, "</FindItem>")
	require.GreaterOrEqual(t, j, 0, soap)
	return soap[i : j+len("</FindItem>")]
}

func TestFindItemsPageRangeRequestGolden(t *testing.T) {
	srv, _, bodies := newServer(t, reply{200, fixture(t, "find_page1.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	from := time.Date(2024, 5, 1, 0, 0, 0, 0, time.UTC)
	to := time.Date(2024, 5, 3, 2, 0, 0, 0, time.FixedZone("x", 2*3600)) // 2024-05-03T00:00:00Z
	_, err := FindItemsPage(context.Background(), c, FolderRef{DistinguishedId: "inbox"}, FindItemsOptions{
		Restriction: DateTimeReceivedRange(from, to),
		MaxEntries:  2,
		SortBy:      "item:DateTimeReceived",
		Ascending:   true,
		Mailbox:     "shared@example.com",
	})
	require.NoError(t, err)

	got := requestBody(t, (*bodies)[0])
	if os.Getenv("UPDATE_GOLDEN") != "" {
		require.NoError(t, os.WriteFile("testdata/find_range_request.golden.xml", []byte(got+"\n"), 0o644))
	}
	assert.Equal(t, strings.TrimSpace(fixture(t, "find_range_request.golden.xml")), got)
}

func TestFindItemsPagePaging(t *testing.T) {
	srv, _, bodies := newServer(t, reply{200, fixture(t, "find_page1.xml")}, reply{200, fixture(t, "find_page2.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	folder := FolderRef{DistinguishedId: "inbox"}
	ctx := context.Background()

	p1, err := FindItemsPage(ctx, c, folder, FindItemsOptions{MaxEntries: 2})
	require.NoError(t, err)
	require.Len(t, p1.Items, 2)
	assert.Equal(t, FoundItem{
		ItemId: "AAMk1", ChangeKey: "CK1", InternetMessageId: "<a@x>",
		DateTimeReceived: time.Date(2024, 5, 1, 9, 0, 0, 0, time.UTC), ParentFolderId: "AAMkInbox==",
	}, p1.Items[0])
	assert.Equal(t, 3, p1.TotalItemsInView)
	assert.False(t, p1.IncludesLastItemInRange)
	assert.Equal(t, 2, p1.NextOffset)
	assert.Equal(t, "V2017_07_11", p1.ServerVersionInfo.Version)

	p2, err := FindItemsPage(ctx, c, folder, FindItemsOptions{MaxEntries: 2, Offset: p1.NextOffset})
	require.NoError(t, err)
	require.Len(t, p2.Items, 1)
	assert.Equal(t, "AAMk3", p2.Items[0].ItemId)
	assert.True(t, p2.IncludesLastItemInRange)
	assert.Equal(t, 3, p2.NextOffset)

	assert.Contains(t, (*bodies)[1], `Offset="2"`)
	assert.Contains(t, (*bodies)[0], `Traversal="Shallow"`)
}

func TestFindItemsPageEmptyWindow(t *testing.T) {
	srv, _, _ := newServer(t, reply{200, fixture(t, "find_empty.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	res, err := FindItemsPage(context.Background(), c, FolderRef{DistinguishedId: "inbox"}, FindItemsOptions{})
	require.NoError(t, err)
	assert.Empty(t, res.Items)
	assert.NotNil(t, res.Items)
	assert.Equal(t, 0, res.TotalItemsInView)
	assert.True(t, res.IncludesLastItemInRange)
}

func TestFindItemsPageServerBusy(t *testing.T) {
	srv, _, _ := newServer(t, reply{200, fixture(t, "find_error_server_busy.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := FindItemsPage(context.Background(), c, FolderRef{DistinguishedId: "inbox"}, FindItemsOptions{})
	var re *ResponseError
	require.True(t, errors.As(err, &re))
	assert.True(t, re.IsServerBusy())
	assert.Equal(t, 45000, re.BackOffMilliseconds)
}

func TestFindItemsPageUnauthorized(t *testing.T) {
	srv, _, _ := newServer(t, reply{401, fixture(t, "unauthorized_401.html")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := FindItemsPage(context.Background(), c, FolderRef{DistinguishedId: "inbox"}, FindItemsOptions{})
	assert.True(t, errors.Is(err, ErrUnauthorized))
}

func TestFindItemsPagePageSizeBounds(t *testing.T) {
	srv, _, bodies := newServer(t, reply{200, fixture(t, "find_empty.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	ctx := context.Background()
	f := FolderRef{Id: "AAMkFolder=="}
	_, err := FindItemsPage(ctx, c, f, FindItemsOptions{})
	require.NoError(t, err)
	_, err = FindItemsPage(ctx, c, f, FindItemsOptions{MaxEntries: 5000})
	require.NoError(t, err)
	assert.Contains(t, (*bodies)[0], `MaxEntriesReturned="100"`)
	assert.Contains(t, (*bodies)[1], `MaxEntriesReturned="1000"`)

	_, err = FindItemsPage(ctx, c, FolderRef{}, FindItemsOptions{})
	assert.Error(t, err)
	_, err = FindItemsPage(ctx, c, f, FindItemsOptions{Offset: -1})
	assert.Error(t, err)
}

// Golden captured from the pre-change code: legacy IsEqualTo output stays byte-identical.
func TestRestrictionIsEqualToUnchanged(t *testing.T) {
	r := NewFindItemRequest(DistinguishedFolderId{Id: "inbox"}, FindItemRequestConfig{Restriction: &Restriction{IsEqualTo: &IsEqualTo{
		FieldURI:           &FieldURI{FieldURI: "item:Subject"},
		FieldURIOrConstant: &FieldURIOrConstant{Constant: &Constant{Value: "x"}},
	}}})
	b, err := xml.MarshalIndent(r, "", "  ")
	require.NoError(t, err)
	assert.Equal(t, fixture(t, "find_isequalto_request.golden.xml"), string(b))
}

// Folder ids and replies below model an Exchange 2019 server: three folders of six
// messages each, received one hour apart and interleaved across the folders.
const (
	probeFolderA = "folder-001"
	probeFolderB = "folder-002"
	probeFolderC = "folder-003"
)

var probeFolders = []FolderRef{{Id: probeFolderA}, {Id: probeFolderB}, {Id: probeFolderC}}

func receivedTimes(items []FoundItem) []string {
	out := make([]string, 0, len(items))
	for _, it := range items {
		out = append(out, it.DateTimeReceived.Format("01-02T15"))
	}
	return out
}

func TestFindItemsPagesRequestGolden(t *testing.T) {
	srv, reqs, bodies := newServer(t, reply{200, fixture(t, "find_pages_asc_off5.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	since := time.Date(2026, 10, 1, 16, 0, 0, 0, time.FixedZone("x", 2*3600)) // 2026-10-01T14:00:00Z
	_, err := FindItemsPages(context.Background(), c, []FolderRef{{Id: "AAOrders=="}, {DistinguishedId: "inbox"}, {Id: "AAArchive==", ChangeKey: "CK"}}, FindItemsOptions{
		Restriction: DateTimeReceivedSince(since),
		Offset:      5,
		MaxEntries:  5,
		SortBy:      "item:DateTimeReceived",
		Mailbox:     "shared@example.com",
	}, WithAnchorMailbox("shared@example.com"))
	require.NoError(t, err)
	assert.Equal(t, "shared@example.com", (*reqs)[0].Header.Get("X-AnchorMailbox"))
	golden(t, "find_pages_request.golden.xml", bodyOf(t, (*bodies)[0], "FindItem"))
}

func TestFindItemsPagesPagesEveryFolderInLockstep(t *testing.T) {
	srv, _, bodies := newServer(t,
		reply{200, fixture(t, "find_pages_asc_off0.xml")},
		reply{200, fixture(t, "find_pages_asc_off5.xml")},
		reply{200, fixture(t, "find_pages_asc_off10.xml")},
	)
	c := NewContextClient(srv.URL, "u", "p", nil)
	ctx := context.Background()
	o := FindItemsOptions{MaxEntries: 5, SortBy: "item:DateTimeReceived", Ascending: true}

	p1, err := FindItemsPages(ctx, c, probeFolders, o)
	require.NoError(t, err)
	require.Len(t, p1, 3)
	for i, want := range [][]string{
		{"10-01T08", "10-01T11", "10-01T14", "10-01T17", "10-01T20"},
		{"10-01T09", "10-01T12", "10-01T15", "10-01T18", "10-01T21"},
		{"10-01T10", "10-01T13", "10-01T16", "10-01T19", "10-01T22"},
	} {
		assert.Equal(t, probeFolders[i], p1[i].Folder)
		require.NoError(t, p1[i].Err)
		assert.Equal(t, want, receivedTimes(p1[i].Page.Items), "folder %d page 1", i)
		assert.Equal(t, probeFolders[i].Id, p1[i].Page.Items[0].ParentFolderId)
		assert.False(t, p1[i].Page.Done)
		assert.Equal(t, 5, p1[i].Page.NextOffset)
		assert.Equal(t, 6, p1[i].Page.TotalItemsInView)
		assert.Equal(t, "V2017_07_11", p1[i].Page.ServerVersionInfo.Version)
	}

	o.Offset = 5
	p2, err := FindItemsPages(ctx, c, probeFolders, o)
	require.NoError(t, err)
	assert.Equal(t, []string{"10-01T23"}, receivedTimes(p2[0].Page.Items))
	assert.Equal(t, []string{"10-02T00"}, receivedTimes(p2[1].Page.Items))
	assert.Equal(t, []string{"10-02T01"}, receivedTimes(p2[2].Page.Items))
	for _, p := range p2 {
		assert.True(t, p.Page.Done, "a short page ends the folder")
		assert.Equal(t, 6, p.Page.NextOffset)
	}

	// Past a folder's end the server echoes the offset as TotalItemsInView.
	o.Offset = 10
	p3, err := FindItemsPages(ctx, c, probeFolders, o)
	require.NoError(t, err)
	for _, p := range p3 {
		assert.Empty(t, p.Page.Items)
		assert.True(t, p.Page.Done)
		assert.Equal(t, 10, p.Page.TotalItemsInView)
	}
	assert.Contains(t, (*bodies)[1], `Offset="5"`)
	assert.Contains(t, (*bodies)[2], `Offset="10"`)
}

func TestFindItemsPagesShortPageWithoutLastItemIsNotDone(t *testing.T) {
	// A server cut short by a throttling policy returns fewer items than asked without IncludesLastItemInRange.
	cut := strings.ReplaceAll(fixture(t, "find_pages_asc_off5.xml"), `IncludesLastItemInRange="true"`, `IncludesLastItemInRange="false"`)
	srv, _, _ := newServer(t, reply{200, cut})
	c := NewContextClient(srv.URL, "u", "p", nil)
	pages, err := FindItemsPages(context.Background(), c, probeFolders, FindItemsOptions{MaxEntries: 5, Offset: 5})
	require.NoError(t, err)
	for _, p := range pages {
		require.Len(t, p.Page.Items, 1)
		assert.False(t, p.Page.Done)
	}
}

func TestFindItemsPagesFollowsTheCallersFolderOrder(t *testing.T) {
	srv, _, bodies := newServer(t, reply{200, fixture(t, "find_pages_mixed.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	// Serialized as inbox, sentitems, A, B, C: distinguished folders go first whatever the input order.
	folders := []FolderRef{{Id: probeFolderA}, {DistinguishedId: "inbox"}, {Id: probeFolderB}, {DistinguishedId: "sentitems"}, {Id: probeFolderC}}
	pages, err := FindItemsPages(context.Background(), c, folders, FindItemsOptions{MaxEntries: 50})
	require.NoError(t, err)
	require.Len(t, pages, 5)
	for i, p := range pages {
		assert.Equal(t, folders[i], p.Folder)
		require.NoError(t, p.Err)
	}
	assert.Equal(t, "10-01T23", receivedTimes(pages[0].Page.Items)[0])
	assert.Equal(t, probeFolderA, pages[0].Page.Items[0].ParentFolderId)
	assert.Len(t, pages[1].Page.Items, 50)
	// The inbox holds exactly 50 messages: a full page whose last item is the folder's last, yet the server
	// says IncludesLastItemInRange=false. It is not Done; the next page comes back empty.
	assert.False(t, pages[1].Page.IncludesLastItemInRange)
	assert.False(t, pages[1].Page.Done)
	assert.Equal(t, "10-02T00", receivedTimes(pages[2].Page.Items)[0])
	assert.Equal(t, probeFolderB, pages[2].Page.Items[0].ParentFolderId)
	assert.Equal(t, []string{"10-23T12", "10-23T11", "10-23T10"}, receivedTimes(pages[3].Page.Items))
	assert.True(t, pages[3].Page.Done)
	assert.Equal(t, "10-02T01", receivedTimes(pages[4].Page.Items)[0])
	assert.Equal(t, probeFolderC, pages[4].Page.Items[0].ParentFolderId)

	assert.Regexp(t, `(?s)Id="inbox".*Id="sentitems".*`+regexp.QuoteMeta(probeFolderA)+`.*`+regexp.QuoteMeta(probeFolderB)+`.*`+regexp.QuoteMeta(probeFolderC), bodyOf(t, (*bodies)[0], "ParentFolderIds"))
}

func TestFindItemsPagesReportsAFolderErrorAndKeepsTheOthers(t *testing.T) {
	srv, _, _ := newServer(t, reply{200, fixture(t, "find_pages_folder_error.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	pages, err := FindItemsPages(context.Background(), c, probeFolders, FindItemsOptions{MaxEntries: 5})
	require.NoError(t, err)
	require.Len(t, pages, 3)

	var re *ResponseError
	require.True(t, errors.As(pages[1].Err, &re))
	assert.Equal(t, "ErrorFolderNotFound", re.Code)
	assert.Nil(t, pages[1].Page)

	require.NoError(t, pages[0].Err)
	assert.Equal(t, "10-01T08", receivedTimes(pages[0].Page.Items)[0])
	require.NoError(t, pages[2].Err)
	assert.Equal(t, "10-01T10", receivedTimes(pages[2].Page.Items)[0])
}

func TestFindItemsPagesKeepsEveryFolderErrorWhenEveryFolderFailed(t *testing.T) {
	srv, _, _ := newServer(t, reply{200, fixture(t, "find_pages_all_failed.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	pages, err := FindItemsPages(context.Background(), c, probeFolders, FindItemsOptions{MaxEntries: 5})
	require.NoError(t, err)
	require.Len(t, pages, 3)
	for i, want := range []string{"ErrorFolderNotFound", "ErrorFolderNotFound", "ErrorAccessDenied"} {
		var re *ResponseError
		require.True(t, errors.As(pages[i].Err, &re), pages[i].Err)
		assert.Equal(t, want, re.Code)
		assert.Nil(t, pages[i].Page)
	}
}

func TestFindItemsPagesReturnsARequestLevelError(t *testing.T) {
	srv, _, _ := newServer(t, reply{200, fixture(t, "find_error_server_busy.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := FindItemsPages(context.Background(), c, probeFolders, FindItemsOptions{})
	var re *ResponseError
	require.True(t, errors.As(err, &re), err)
	assert.True(t, re.IsServerBusy())
	assert.Equal(t, 45000, re.BackOffMilliseconds)
}

func TestFindItemsPagesRejectsAResponseCountMismatch(t *testing.T) {
	srv, _, _ := newServer(t, reply{200, fixture(t, "find_pages_asc_off5.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := FindItemsPages(context.Background(), c, probeFolders[:2], FindItemsOptions{})
	require.Error(t, err)
	assert.Contains(t, err.Error(), "asked 2 folders, got 3 response messages")
}

func TestFindItemsPagesValidatesItsInput(t *testing.T) {
	c := NewContextClient("http://unused", "u", "p", nil)
	ctx := context.Background()
	_, err := FindItemsPages(ctx, c, nil, FindItemsOptions{})
	require.Error(t, err)
	assert.Contains(t, err.Error(), "needs a folder")
	_, err = FindItemsPages(ctx, c, []FolderRef{{Id: "AAx"}, {}}, FindItemsOptions{})
	require.Error(t, err)
	assert.Contains(t, err.Error(), "needs non-empty folders")
	_, err = FindItemsPages(ctx, c, probeFolders, FindItemsOptions{Offset: -1})
	require.Error(t, err)
	assert.Contains(t, err.Error(), "offset must not be negative")
}
