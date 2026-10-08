package ewsutil

import (
	"context"

	"github.com/hoshii-ai/ews"
	"github.com/pkg/errors"
)

// GetInboxCategoriesContext is GetInboxCategories for an explicit mailbox. mailbox
// is the SMTP address whose category list is read (e.g. a delegate mailbox); it
// is not taken from the client's login. Pair it with ews.WithAnchorMailbox.
func GetInboxCategoriesContext(ctx context.Context, c ews.ContextClient, mailbox string, opts ...ews.RequestOption) (*ews.CategoryList, error) {
	// MS Exchange stores categories in a configuration item of the calendar folder
	found, err := ews.FindItemsPage(ctx, c, ews.FolderRef{DistinguishedId: "calendar"}, ews.FindItemsOptions{
		Associated:           true,
		Mailbox:              mailbox,
		AdditionalProperties: []string{"item:ItemClass"},
		Restriction: &ews.Restriction{IsEqualTo: &ews.IsEqualTo{
			FieldURI:           &ews.FieldURI{FieldURI: "item:ItemClass"},
			FieldURIOrConstant: &ews.FieldURIOrConstant{Constant: &ews.Constant{Value: "IPM.Configuration.CategoryList"}},
		}},
	}, opts...)
	if err != nil {
		return nil, errors.Wrap(err, "failed to find category list item")
	}
	if len(found.Items) == 0 {
		return nil, errors.New("category list item not found")
	}
	if len(found.Items) > 1 {
		return nil, errors.Errorf("expected 1 category list item, got %d", len(found.Items))
	}
	itemId := ews.ItemId{Id: found.Items[0].ItemId, ChangeKey: found.Items[0].ChangeKey}

	resp, err := ews.GetItemContext(ctx, c, itemId, ews.GetItemRequestConfig{
		ItemShape: &ews.ItemShape{
			BaseShape: ews.BaseShapeAllProperties,
			AdditionalProperties: &ews.AdditionalProperties{
				ExtendedFieldURI: []ews.ExtendedFieldURI{{PropertyTag: ews.PropertyTagCategories, PropertyType: ews.PropertyTypeBinary}},
			},
		},
	}, opts...)
	if err != nil {
		return nil, errors.Wrap(err, "failed to get category list item")
	}

	messages := resp.ResponseMessages.GetItemResponseMessage.Items.Message
	if len(messages) != 1 {
		return nil, errors.Errorf("expected 1 message, got %d", len(messages))
	}
	message := messages[0]
	props := message.ExtendedProperties
	if len(props) != 1 {
		return nil, errors.Errorf("expected 1 extended property, got %d", len(props))
	}
	if props[0].ExtendedFieldURI.PropertyTag != ews.PropertyTagCategories {
		return nil, errors.Errorf("expected property tag categories, got %s", props[0].ExtendedFieldURI.PropertyTag)
	}
	if props[0].Value == nil {
		return nil, errors.New("extended property value is nil")
	}

	list, err := ews.CategoryListFromBase64(*props[0].Value)
	if err != nil {
		return nil, errors.Wrap(err, "failed to decode categories")
	}
	list.ItemId = itemId
	if message.ItemId != nil {
		list.ItemId = *message.ItemId
	}
	return list, nil
}

// AddCategoriesContext is AddCategories for an explicit mailbox. Categories whose
// name already exists are skipped; nothing is written if nothing changed.
func AddCategoriesContext(ctx context.Context, c ews.ContextClient, mailbox string, categories []ews.Category, opts ...ews.RequestOption) error {
	list, err := GetInboxCategoriesContext(ctx, c, mailbox, opts...)
	if err != nil {
		return errors.Wrap(err, "failed to get inbox categories")
	}
	oldData, err := list.CategoryListToBase64()
	if err != nil {
		return errors.Wrap(err, "failed to convert category list to base64")
	}
	for _, category := range categories {
		if err := list.AddCategory(category.Name, category.Color); err != nil {
			continue // duplicate
		}
	}
	data, err := list.CategoryListToBase64()
	if err != nil {
		return errors.Wrap(err, "failed to convert category list to base64")
	}
	if data == oldData {
		return nil
	}

	_, err = ews.UpdateItemContext(ctx, c, &ews.UpdateItemRequest{
		MessageDisposition: ews.MessageDispositionSaveOnly,
		ItemChanges: ews.ItemChanges{ItemChange: []ews.ItemChange{{
			ItemId: list.ItemId,
			Updates: ews.Updates{SetItemField: []ews.SetItemField{{
				ExtendedFieldURI: &ews.ExtendedFieldURI{PropertyTag: ews.PropertyTagCategories, PropertyType: ews.PropertyTypeBinary},
				Message:          &ews.Message{ExtendedProperties: []ews.ExtendedProperty{{Value: &data}}},
			}}},
		}}},
	}, opts...)
	if err != nil {
		return errors.Wrap(err, "failed to update category list item")
	}
	return nil
}
