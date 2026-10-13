package ewsutil

import (
	"context"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/http/httptest"
	"sync"
	"testing"

	"github.com/hoshii-ai/ews"
	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

const errReply = `<?xml version="1.0" encoding="utf-8"?><s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><m:%sResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"><m:ResponseMessages><m:%sResponseMessage ResponseClass="Error"><m:MessageText>nope</m:MessageText><m:ResponseCode>ErrorItemNotFound</m:ResponseCode></m:%sResponseMessage></m:ResponseMessages></m:%sResponse></s:Body></s:Envelope>`

func replyServer(t *testing.T, replies ...string) (ews.ContextClient, *[]string, *[]string) {
	t.Helper()
	var mu sync.Mutex
	var bodies, anchors []string
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		b, _ := io.ReadAll(r.Body)
		mu.Lock()
		i := len(bodies)
		bodies = append(bodies, string(b))
		anchors = append(anchors, r.Header.Get("X-AnchorMailbox"))
		mu.Unlock()
		if i >= len(replies) {
			i = len(replies) - 1
		}
		_, _ = w.Write([]byte(replies[i]))
	}))
	t.Cleanup(srv.Close)
	return ews.NewContextClient(srv.URL, "u", "p", nil), &bodies, &anchors
}

func errResp(op string) string {
	return fmt.Sprintf(errReply, op, op, op, op)
}

func TestSendEmailContext(t *testing.T) {
	c, bodies, anchors := replyServer(t,
		soap("CreateItem", `<m:Items><t:Message><t:ItemId Id="AAdraft" ChangeKey="CK1"/></t:Message></m:Items>`),
		soap("SendItem", ``),
	)
	id, err := SendEmailContext(context.Background(), c, []string{"a@example.com"}, []string{"b@example.com"}, nil, "hi", "<p>x</p>", nil, ews.WithAnchorMailbox("me@example.com"))
	require.NoError(t, err)
	assert.Equal(t, &ews.ItemId{Id: "AAdraft", ChangeKey: "CK1"}, id)
	require.Len(t, *bodies, 2)
	assert.Contains(t, (*bodies)[0], `MessageDisposition="SaveOnly"`)
	assert.Contains(t, (*bodies)[0], "a@example.com")
	assert.Contains(t, (*bodies)[1], `Id="AAdraft"`)
	assert.Equal(t, []string{"me@example.com", "me@example.com"}, *anchors)

	_, err = SendEmailContext(context.Background(), c, nil, nil, nil, "hi", "x", nil)
	assert.Error(t, err)
}

func TestSendEmailContextError(t *testing.T) {
	c, _, _ := replyServer(t, errResp("CreateItem"))
	_, err := SendEmailContext(context.Background(), c, []string{"a@example.com"}, nil, nil, "hi", "x", nil)
	var re *ews.ResponseError
	require.True(t, errors.As(err, &re))
	assert.Equal(t, "ErrorItemNotFound", re.Code)
}

func TestGetMessageContext(t *testing.T) {
	c, bodies, _ := replyServer(t,
		soap("GetItem", `<m:Items><t:Message><t:ItemId Id="AA1" ChangeKey="CK1"/><t:Subject>hello</t:Subject></t:Message></m:Items>`),
		errResp("GetItem"),
	)
	msg, err := GetMessageContext(context.Background(), c, &ews.ItemId{Id: "AA1"})
	require.NoError(t, err)
	assert.Equal(t, "hello", *msg.Subject)
	assert.Contains(t, (*bodies)[0], "item:InternetMessageHeaders")

	_, err = GetMessageContext(context.Background(), c, &ews.ItemId{Id: "AA1"})
	var re *ews.ResponseError
	require.True(t, errors.As(err, &re))

	_, err = GetMessageContext(context.Background(), c, nil)
	assert.Error(t, err)
}

func TestGetMessageByInternetMessageIdContext(t *testing.T) {
	find := soap("FindItem", `<m:RootFolder TotalItemsInView="1" IncludesLastItemInRange="true" IndexedPagingOffset="1"><t:Items><t:Message><t:ItemId Id="AA1" ChangeKey="CK1"/><t:InternetMessageId>&lt;a@x&gt;</t:InternetMessageId></t:Message></t:Items></m:RootFolder>`)
	get := soap("GetItem", `<m:Items><t:Message><t:ItemId Id="AA1" ChangeKey="CK1"/><t:Subject>hello</t:Subject></t:Message></m:Items>`)
	empty := soap("FindItem", `<m:RootFolder TotalItemsInView="0" IncludesLastItemInRange="true" IndexedPagingOffset="0"><t:Items/></m:RootFolder>`)
	c, bodies, anchors := replyServer(t, find, get, empty)

	msg, err := GetMessageByInternetMessageIdContext(context.Background(), c, "<a@x>", ews.WithAnchorMailbox("me@example.com"))
	require.NoError(t, err)
	assert.Equal(t, "hello", *msg.Subject)
	assert.Contains(t, (*bodies)[0], `Id="inbox"`)
	assert.Equal(t, []string{"me@example.com", "me@example.com"}, *anchors)

	_, err = GetMessageByInternetMessageIdContext(context.Background(), c, "<none@x>")
	assert.Error(t, err)
}

func TestUpdateEmailCategoriesContext(t *testing.T) {
	c, bodies, anchors := replyServer(t, soap("UpdateItem", ``), errResp("UpdateItem"))
	id := &ews.ItemId{Id: "AA1", ChangeKey: "CK1"}
	got, err := UpdateEmailCategoriesContext(context.Background(), c, id, []string{"Red", "Blue"}, ews.WithAnchorMailbox("me@example.com"))
	require.NoError(t, err)
	assert.Equal(t, id, got)
	assert.Contains(t, (*bodies)[0], "item:Categories")
	assert.Contains(t, (*bodies)[0], "Red")
	assert.Equal(t, "me@example.com", (*anchors)[0])

	_, err = UpdateEmailCategoriesContext(context.Background(), c, id, nil)
	var re *ews.ResponseError
	require.True(t, errors.As(err, &re))
	assert.Equal(t, "ErrorItemNotFound", re.Code)

	_, err = UpdateEmailCategoriesContext(context.Background(), c, nil, nil)
	assert.Error(t, err)
}
