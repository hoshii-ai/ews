package ewsutil

import (
	"context"
	"fmt"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/hoshii-ai/ews"
	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

const (
	soapOpen  = `<?xml version="1.0" encoding="utf-8"?><s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><m:%sResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"><m:ResponseMessages><m:%sResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode>`
	soapClose = `</m:%sResponseMessage></m:ResponseMessages></m:%sResponse></s:Body></s:Envelope>`
)

func soap(op, inner string) string {
	return fmt.Sprintf(soapOpen, op, op) + inner + fmt.Sprintf(soapClose, op, op)
}

func TestCategoriesForMailbox(t *testing.T) {
	list := &ews.CategoryList{Default: "x", LastSavedSession: 1, LastSavedTime: time.Now().UTC()}
	require.NoError(t, list.AddCategory("Existing", ews.ColorBlue))
	b64, err := list.CategoryListToBase64()
	require.NoError(t, err)

	replies := []string{
		soap("FindItem", `<m:RootFolder TotalItemsInView="1" IncludesLastItemInRange="true" IndexedPagingOffset="1"><t:Items><t:Message><t:ItemId Id="AAcfg" ChangeKey="CK1"/></t:Message></t:Items></m:RootFolder>`),
		soap("GetItem", `<m:Items><t:Message><t:ItemId Id="AAcfg" ChangeKey="CK1"/><t:ExtendedProperty><t:ExtendedFieldURI PropertyTag="0x7c08" PropertyType="Binary"/><t:Value>`+b64+`</t:Value></t:ExtendedProperty></t:Message></m:Items>`),
		soap("UpdateItem", `<m:Items><t:Message><t:ItemId Id="AAcfg" ChangeKey="CK2"/></t:Message></m:Items>`),
	}
	var mu sync.Mutex
	var bodies, anchors []string
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		b, _ := io.ReadAll(r.Body)
		mu.Lock()
		i := len(bodies)
		bodies = append(bodies, string(b))
		anchors = append(anchors, r.Header.Get("X-AnchorMailbox"))
		mu.Unlock()
		_, _ = w.Write([]byte(replies[i]))
	}))
	defer srv.Close()

	c := ews.NewContextClient(srv.URL, "svc-login", "p", nil)
	err = AddCategoriesContext(context.Background(), c, "shared@example.com",
		[]ews.Category{{Name: "Existing", Color: ews.ColorRed}, {Name: "New", Color: ews.ColorGreen}},
		ews.WithAnchorMailbox("shared@example.com"))
	require.NoError(t, err)

	require.Len(t, bodies, 3)
	// the mailbox comes from the argument, not the login
	assert.Contains(t, bodies[0], `<EmailAddress xmlns="http://schemas.microsoft.com/exchange/services/2006/types">shared@example.com</EmailAddress>`)
	assert.NotContains(t, bodies[0], "svc-login")
	assert.Contains(t, bodies[0], `Traversal="Associated"`)
	assert.Contains(t, bodies[0], "IPM.Configuration.CategoryList")
	assert.Contains(t, bodies[1], `PropertyTag="0x7c08"`)
	assert.Contains(t, bodies[2], "<UpdateItem")
	assert.True(t, strings.Contains(bodies[2], `Id="AAcfg"`))
	for _, a := range anchors {
		assert.Equal(t, "shared@example.com", a)
	}
}

func TestCategoriesNoChangeSkipsUpdate(t *testing.T) {
	list := &ews.CategoryList{Default: "x", LastSavedSession: 1, LastSavedTime: time.Now().UTC()}
	require.NoError(t, list.AddCategory("Existing", ews.ColorBlue))
	b64, _ := list.CategoryListToBase64()
	replies := []string{
		soap("FindItem", `<m:RootFolder TotalItemsInView="1" IncludesLastItemInRange="true"><t:Items><t:Message><t:ItemId Id="AAcfg" ChangeKey="CK1"/></t:Message></t:Items></m:RootFolder>`),
		soap("GetItem", `<m:Items><t:Message><t:ItemId Id="AAcfg"/><t:ExtendedProperty><t:ExtendedFieldURI PropertyTag="0x7c08" PropertyType="Binary"/><t:Value>`+b64+`</t:Value></t:ExtendedProperty></t:Message></m:Items>`),
	}
	n := 0
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		_, _ = w.Write([]byte(replies[n]))
		n++
	}))
	defer srv.Close()
	c := ews.NewContextClient(srv.URL, "u", "p", nil)
	require.NoError(t, AddCategoriesContext(context.Background(), c, "a@example.com", []ews.Category{{Name: "Existing"}}))
	assert.Equal(t, 2, n)
}
