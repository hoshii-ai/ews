package ewsutil

import (
	"testing"

	"github.com/hoshii-ai/ews"
	"github.com/hoshii-ai/ews/utils"
	"github.com/stretchr/testify/require"
)

// CreateEmailDraft helper method to create email draft
func Test_CreateEmailDraft(t *testing.T) {
	client := ews.NewClient(url, username, password, &ews.Config{
		Dump:    true,
		NTLM:    true,
		SkipTLS: false,
	})

	attachments := []ews.FileAttachment{
		{
			Name:           "test.jpeg",
			IsInline:       utils.Ptr(false),
			ContentType:    "image/jpeg",
			IsContactPhoto: utils.Ptr(false),
			Content:        "dGVzdA==",
		},
	}
	itemId, err := CreateEmailDraft(client, []string{"test@outlook.com"}, "Test Subject 2", "Test Body", attachments...)
	if err != nil {
		t.Fatalf("failed to create email draft: %v", err)
	}

	t.Log("email draft created successfully", itemId)
}

func Test_SendEmailDraft(t *testing.T) {
	client := ews.NewClient(url, username, password, &ews.Config{
		Dump:    true,
		NTLM:    true,
		SkipTLS: false,
	})
	attachments := []ews.FileAttachment{
		{
			Name:           "test.jpeg",
			IsInline:       utils.Ptr(false),
			ContentType:    "image/jpeg",
			IsContactPhoto: utils.Ptr(false),
			Content:        "dGVzdA==",
		},
	}
	itemId, err := CreateEmailDraft(client, []string{"hoshii.mustermann@gmail.com"}, "Test Subject", "Test Body", attachments...)
	if err != nil {
		t.Fatalf("failed to create email draft: %v", err)
	}

	message, err := ews.GetItem(client, *itemId, ews.GetItemRequestConfig{
		ItemShape: &ews.ItemShape{
			BaseShape: ews.BaseShapeAllProperties,
		},
	})

	t.Logf("internet message id: %+v", *message.ResponseMessages.GetItemResponseMessage.Items.Message[0].InternetMessageId)

	if err != nil {
		t.Fatalf("failed to get item: %v", err)
	}

	t.Logf("message: %+v", message)

	// send the email draft
	sendItemResponse, err := ews.SendItem(client, *itemId, true)
	if err != nil {
		t.Fatalf("failed to send email draft: %v", err)
	}

	t.Logf("send item response: %+v", sendItemResponse)

	require.Equal(t, ews.ResponseClassSuccess, sendItemResponse.ResponseMessages.SendItemResponseMessage.ResponseClass)
}
