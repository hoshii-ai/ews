package ewsutil

import (
	"github.com/hoshii-ai/ews"
	"github.com/hoshii-ai/ews/utils"
	"github.com/pkg/errors"
)

func GetInboxCategories(c ews.Client) (*ews.CategoryList, error) {
	// MS Exchange stores categories in the calendar folder
	findItemConfig := ews.FindItemRequestConfig{
		Traversal: utils.Ptr(ews.FindItemTraversalAssociated),
		BaseShape: utils.Ptr(ews.BaseShapeIdOnly),
		AdditionalProperties: &ews.AdditionalProperties{
			FieldURI: []ews.FieldURI{
				{
					FieldURI: "item:ItemClass",
				},
			},
		},
		Restriction: &ews.Restriction{
			IsEqualTo: &ews.IsEqualTo{
				FieldURI: &ews.FieldURI{
					FieldURI: "item:ItemClass",
				},
				FieldURIOrConstant: &ews.FieldURIOrConstant{
					Constant: &ews.Constant{
						Value: "IPM.Configuration.CategoryList",
					},
				},
			},
		},
	}

	findItemResponse, err := ews.FindItem(c, "calendar", findItemConfig)
	if err != nil {
		return nil, errors.Wrap(err, "failed to find item")
	}

	if findItemResponse.ResponseMessages.FindItemResponseMessage.ResponseClass != ews.ResponseClassSuccess {
		return nil, errors.New("failed to find item: " + findItemResponse.ResponseMessages.FindItemResponseMessage.ResponseCode)
	}

	rootFolder := findItemResponse.ResponseMessages.FindItemResponseMessage.RootFolder

	messages := rootFolder.Items.Message
	if len(messages) == 0 {
		return nil, errors.New("no messages found in rootFolder.Items")
	}
	if len(messages) > 1 {
		return nil, errors.Errorf("expected 1 messages, got %d", len(messages))
	}

	message := messages[0]

	if message.ItemId == nil {
		return nil, errors.New("message item id is nil")
	}

	getItemConfig := ews.GetItemRequestConfig{
		ItemShape: &ews.ItemShape{
			BaseShape: ews.BaseShapeAllProperties,
			AdditionalProperties: &ews.AdditionalProperties{
				ExtendedFieldURI: []ews.ExtendedFieldURI{
					{
						PropertyTag:  ews.PropertyTagCategories,
						PropertyType: ews.PropertyTypeBinary,
					},
				},
			},
		},
	}

	getItemResponse, err := ews.GetItem(c, *message.ItemId, getItemConfig)
	if err != nil {
		return nil, errors.Wrap(err, "failed to get item")
	}

	if getItemResponse.ResponseMessages.GetItemResponseMessage.ResponseClass != ews.ResponseClassSuccess {
		return nil, errors.New("failed to get item: " + getItemResponse.ResponseMessages.GetItemResponseMessage.ResponseCode)
	}

	messages = getItemResponse.ResponseMessages.GetItemResponseMessage.Items.Message
	if len(messages) != 1 {
		return nil, errors.Errorf("expected 1 message, got %d", len(messages))
	}
	message = messages[0]

	extendedProperties := message.ExtendedProperties
	if len(extendedProperties) != 1 {
		return nil, errors.Errorf("expected 1 extended property, got %d", len(extendedProperties))
	}
	if extendedProperties[0].ExtendedFieldURI.PropertyTag != ews.PropertyTagCategories {
		return nil, errors.Errorf("expected property tag categories, got %s", extendedProperties[0].ExtendedFieldURI.PropertyTag)
	}
	if extendedProperties[0].ExtendedFieldURI.PropertyType != ews.PropertyTypeBinary {
		return nil, errors.Errorf("expected property type binary, got %s", extendedProperties[0].ExtendedFieldURI.PropertyType)
	}
	if extendedProperties[0].Value == nil {
		return nil, errors.New("extended property value is nil")
	}

	categories, err := ews.CategoryListFromBase64(*extendedProperties[0].Value)
	if err != nil {
		return nil, errors.Wrap(err, "failed to decode categories")
	}
	// Safe: checked above
	categories.ItemId = *message.ItemId

	return categories, nil
}
