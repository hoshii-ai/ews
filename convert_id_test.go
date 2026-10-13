package ews

import (
	"context"
	"errors"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestConvertId(t *testing.T) {
	srv, reqs, bodies := newServer(t, reply{200, fixture(t, "convert_id_ok.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	res, err := ConvertId(context.Background(), c, []AlternateId{
		{Format: IdFormatEwsId, Id: "AQMkOne==", Mailbox: "a@example.com"},
		{Format: IdFormatEwsId, Id: "bad", Mailbox: "a@example.com"},
	}, IdFormatHexEntryId, WithAnchorMailbox("a@example.com"))
	require.NoError(t, err)
	require.Len(t, res, 2)
	assert.Equal(t, AlternateId{Format: IdFormatEwsId, Id: "AAMkOne==", Mailbox: "a@example.com"}, res[0].Id)
	assert.NoError(t, res[0].Err)
	var re *ResponseError
	require.True(t, errors.As(res[1].Err, &re))
	assert.Equal(t, "ErrorInvalidIdMalformed", re.Code)
	assert.Equal(t, "a@example.com", (*reqs)[0].Header.Get("X-AnchorMailbox"))
	golden(t, "convert_id_request.golden.xml", bodyOf(t, (*bodies)[0], "ConvertId"))
}

func TestConvertIdRequestLevelError(t *testing.T) {
	srv, _, _ := newServer(t, reply{200, fixture(t, "convert_id_busy.xml")})
	c := NewContextClient(srv.URL, "u", "p", nil)
	_, err := ConvertId(context.Background(), c, []AlternateId{
		{Format: IdFormatEwsId, Id: "a", Mailbox: "m"}, {Format: IdFormatEwsId, Id: "b", Mailbox: "m"},
	}, IdFormatEwsId)
	var re *ResponseError
	require.True(t, errors.As(err, &re))
	assert.True(t, re.IsServerBusy())
	assert.Equal(t, 1000, re.BackOffMilliseconds)
}

func TestConvertIdValidatesInput(t *testing.T) {
	c := NewContextClient("http://unused", "u", "p", nil)
	ctx := context.Background()
	ok := AlternateId{Format: IdFormatEwsId, Id: "a", Mailbox: "m"}
	_, err := ConvertId(ctx, c, nil, IdFormatEwsId)
	assert.Error(t, err)
	_, err = ConvertId(ctx, c, []AlternateId{ok}, "Nope")
	assert.Error(t, err)
	_, err = ConvertId(ctx, c, []AlternateId{{Format: IdFormatEwsId}}, IdFormatEwsId)
	assert.Error(t, err)
	_, err = ConvertId(ctx, c, []AlternateId{{Id: "a"}}, IdFormatEwsId)
	assert.Error(t, err)
}
