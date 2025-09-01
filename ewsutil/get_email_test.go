package ewsutil

import (
	"testing"

	"github.com/hoshii-ai/ews"
	"github.com/stretchr/testify/require"
)

func Test_GetMessageByInternetMessageId(t *testing.T) {
	c := ews.NewClient(url, username, password, &ews.Config{
		Dump:    true,
		NTLM:    true,
		SkipTLS: false,
	})

	message, err := GetMessageByInternetMessageId(c, "<CAKMBqB1yyODh+Bjpv9bL7attSoKxPo-a03t_N5QKKt_+eBPh1w@mail.gmail.com>")
	if err != nil {
		t.Fatalf("failed to get message: %v", err)
	}

	t.Logf("message: %+v", *message.Subject)

	require.Equal(t, "<CAKMBqB1yyODh+Bjpv9bL7attSoKxPo-a03t_N5QKKt_+eBPh1w@mail.gmail.com>", *message.InternetMessageId)

	require.Greater(t, len(message.Categories.String), 0)
	for _, category := range message.Categories.String {
		t.Logf("category: %s", category)
	}
}
