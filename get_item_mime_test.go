package ews

import (
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestGetItemMimeContent(t *testing.T) {
	srv, _, bodies := newServer(t, reply{200, fixture(t, "get_item_mime.xml")})
	c := NewClient(srv.URL, "u", "p", &Config{})
	resp, err := GetItem(c, ItemId{Id: "AAMk1"}, GetItemRequestConfig{ItemShape: NewGetItemMimeShape()})
	require.NoError(t, err)
	assert.Contains(t, (*bodies)[0], ">true</IncludeMimeContent>")

	m := resp.ResponseMessages.GetItemResponseMessage.Items.Message[0]
	require.NotNil(t, m.MimeContent)
	assert.Equal(t, "UTF-8", m.MimeContent.CharacterSet)
	b, err := m.MimeContent.Decode()
	require.NoError(t, err)
	assert.Equal(t, "Subject: hi\r\n\r\nbody", string(b))
}
