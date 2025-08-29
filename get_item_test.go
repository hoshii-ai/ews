package ews

import (
	"fmt"
	"os"
	"testing"
)

var (
	url      = os.Getenv("EWS_URL")
	username = os.Getenv("EWS_USERNAME")
	password = os.Getenv("EWS_PASSWORD")
)

func TestGetItem(t *testing.T) {
	client := NewClient(url, username, password, &Config{
		Dump: true,
		NTLM: true,
	})

	item, err := GetItem(client,
		ItemId{
			Id:        "AAMkADMxNGE3NDBlLWM5YTgtNDcxNC1hY2NlLTU4NjM0MjgyYzJmYgBGAAAAAADOjREMuZ5MSpIvu32JDVT2BwDtfU8UYrr4T4BgrN2SNR99AAAAAAEMAADtfU8UYrr4T4BgrN2SNR99AAKYV3DaAAA=",
			ChangeKey: "CQAAABYAAADtfU8UYrr4T4BgrN2SNR99AAKYV20B",
		},
		GetItemRequestConfig{
			ItemShape: &ItemShape{
				BaseShape: BaseShapeAllProperties,
				AdditionalProperties: &AdditionalProperties{
					FieldURI: []FieldURI{
						{FieldURI: "item:Attachments"},
					},
				},
			},
		},
	)
	if err != nil {
		t.Fatalf("failed to get item: %v", err)
	}

	for _, header := range item.ResponseMessages.GetItemResponseMessage.Items.Message[0].InternetMessageHeaders.InternetMessageHeader {
		fmt.Println(header.HeaderName + ": " + header.Value)
	}
	fmt.Println("internet message id: " + *item.ResponseMessages.GetItemResponseMessage.Items.Message[0].InternetMessageId)

	for _, attachment := range item.ResponseMessages.GetItemResponseMessage.Items.Message[0].Attachments.FileAttachment {
		fmt.Println("attachment: " + attachment.Name)
	}

	for _, category := range item.ResponseMessages.GetItemResponseMessage.Items.Message[0].Categories.String {
		fmt.Println("category: " + category)
	}
}
