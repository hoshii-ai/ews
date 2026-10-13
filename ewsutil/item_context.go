package ewsutil

import (
	"context"

	"github.com/hoshii-ai/ews"
	"github.com/hoshii-ai/ews/utils"
	"github.com/pkg/errors"
)

// SendEmailContext is SendEmail with a context and per-request options (e.g.
// ews.WithAnchorMailbox). The message is saved to drafts, then sent. A
// ResponseClass=Error reply is returned (wrapped) as *ews.ResponseError.
func SendEmailContext(
	ctx context.Context,
	c ews.ContextClient,
	to []string,
	cc []string,
	bcc []string,
	subject,
	body string,
	attachments []ews.FileAttachment,
	opts ...ews.RequestOption) (*ews.ItemId, error) {
	m, err := buildEmailMessage(to, cc, bcc, subject, body, attachments)
	if err != nil {
		return nil, err
	}

	itemId, err := ews.CreateMessageItemContext(ctx, c, *m, ews.CreateItemRequestConfig{
		MessageDisposition: ews.MessageDispositionSaveOnly,
		SavedItemFolderId:  &ews.SavedItemFolderId{DistinguishedFolderId: ews.DistinguishedFolderId{Id: "drafts"}},
	}, opts...)
	if err != nil {
		return nil, errors.Wrap(err, "failed to send email: failed to create message item")
	}
	if itemId == nil {
		return nil, errors.New("failed to send email: no item id returned for draft")
	}

	if err := ews.SendItemContext(ctx, c, *itemId, true, opts...); err != nil {
		return nil, errors.Wrap(err, "failed to send email")
	}
	return itemId, nil
}

func messageShape() ews.GetItemRequestConfig {
	return ews.GetItemRequestConfig{
		ItemShape: &ews.ItemShape{
			BaseShape: ews.BaseShapeAllProperties,
			AdditionalProperties: &ews.AdditionalProperties{
				FieldURI: []ews.FieldURI{{FieldURI: "item:InternetMessageHeaders"}},
			},
		},
	}
}

// GetMessageContext is GetMessage with a context and per-request options.
func GetMessageContext(ctx context.Context, c ews.ContextClient, itemId *ews.ItemId, opts ...ews.RequestOption) (*ews.Message, error) {
	if itemId == nil {
		return nil, errors.New("item id is required")
	}
	resp, err := ews.GetItemContext(ctx, c, *itemId, messageShape(), opts...)
	if err != nil {
		return nil, errors.Wrap(err, "failed to get item")
	}
	messages := resp.ResponseMessages.GetItemResponseMessage.Items.Message
	if len(messages) != 1 {
		return nil, errors.Errorf("expected 1 message, got %d", len(messages))
	}
	return &messages[0], nil
}

// GetMessageByInternetMessageIdContext is GetMessageByInternetMessageId with a
// context and per-request options. It searches the inbox and expects exactly
// one match.
func GetMessageByInternetMessageIdContext(ctx context.Context, c ews.ContextClient, internetMessageId string, opts ...ews.RequestOption) (*ews.Message, error) {
	found, err := ews.FindItemByInternetMessageID(ctx, c, ews.FolderRef{DistinguishedId: "inbox"}, internetMessageId, opts...)
	if err != nil {
		return nil, errors.Wrap(err, "failed to find item")
	}
	if len(found) != 1 {
		return nil, errors.Errorf("expected 1 message, got %d", len(found))
	}
	return GetMessageContext(ctx, c, &ews.ItemId{Id: found[0].ItemId, ChangeKey: found[0].ChangeKey}, opts...)
}

// UpdateEmailCategoriesContext is UpdateEmailCategories with a context and
// per-request options. It overwrites the existing categories.
func UpdateEmailCategoriesContext(ctx context.Context, c ews.ContextClient, itemId *ews.ItemId, categories []string, opts ...ews.RequestOption) (*ews.ItemId, error) {
	if itemId == nil {
		return nil, errors.New("item id is required")
	}
	req := ews.UpdateItemRequest{
		MessageDisposition: ews.MessageDispositionSaveOnly,
		ConflictResolution: utils.Ptr(ews.ConflictResolutionAlwaysOverwrite),
		ItemChanges: ews.ItemChanges{
			ItemChange: []ews.ItemChange{{
				ItemId: *itemId,
				Updates: ews.Updates{
					SetItemField: []ews.SetItemField{{
						FieldURI: &ews.FieldURI{FieldURI: "item:Categories"},
						Message:  &ews.Message{Categories: &ews.Categories{String: categories}},
					}},
				},
			}},
		},
	}
	if _, err := ews.UpdateItemContext(ctx, c, &req, opts...); err != nil {
		return nil, errors.Wrap(err, "failed to update item")
	}
	return itemId, nil
}
