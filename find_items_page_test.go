package ews

import (
	"context"
	"encoding/xml"
	"errors"
	"os"
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
