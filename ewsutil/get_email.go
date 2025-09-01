package ewsutil

import (
	"strconv"

	"github.com/hoshii-ai/ews"
	"github.com/hoshii-ai/ews/utils"
	"github.com/pkg/errors"
)

func GetMessage(c ews.Client, itemId *ews.ItemId) (*ews.Message, error) {
	getItemConfig := ews.GetItemRequestConfig{
		ItemShape: &ews.ItemShape{
			BaseShape: ews.BaseShapeAllProperties,
			AdditionalProperties: &ews.AdditionalProperties{
				FieldURI: []ews.FieldURI{
					{
						FieldURI: "item:InternetMessageHeaders",
					},
				},
			},
		},
	}
	getItemResponse, err := ews.GetItem(c, *itemId, getItemConfig)
	if err != nil {
		return nil, errors.Wrap(err, "failed to get item")
	}

	if getItemResponse.ResponseMessages.GetItemResponseMessage.ResponseClass != ews.ResponseClassSuccess {
		return nil, errors.New("failed to get item: " + getItemResponse.ResponseMessages.GetItemResponseMessage.ResponseCode)
	}

	messages := getItemResponse.ResponseMessages.GetItemResponseMessage.Items.Message
	if len(messages) != 1 {
		return nil, errors.New("expected 1 message, got " + strconv.Itoa(len(messages)))
	}

	return &messages[0], nil
}

func GetMessageByInternetMessageId(c ews.Client, internetMessageId string) (*ews.Message, error) {
	findItemConfig := ews.FindItemRequestConfig{
		Traversal: utils.Ptr(ews.FindItemTraversalShallow),
		BaseShape: utils.Ptr(ews.BaseShapeIdOnly),
		Restriction: &ews.Restriction{
			IsEqualTo: &ews.IsEqualTo{
				ExtendedFieldURI: &ews.ExtendedFieldURI{
					PropertyTag:  ews.PropertyTagInternetMessageId,
					PropertyType: ews.PropertyTypeString,
				},
				FieldURIOrConstant: &ews.FieldURIOrConstant{
					Constant: &ews.Constant{
						Value: internetMessageId,
					},
				},
			},
		},
	}
	findItemResponse, err := ews.FindItem(c, "inbox", findItemConfig)
	if err != nil {
		return nil, errors.Wrap(err, "failed to find item")
	}

	if findItemResponse.ResponseMessages.FindItemResponseMessage.ResponseClass != ews.ResponseClassSuccess {
		return nil, errors.New("failed to find item: " + findItemResponse.ResponseMessages.FindItemResponseMessage.ResponseCode)
	}

	rootFolder := findItemResponse.ResponseMessages.FindItemResponseMessage.RootFolder

	messages := rootFolder.Items.Message
	if len(messages) != 1 {
		return nil, errors.New("expected 1 message, got " + strconv.Itoa(len(messages)))
	}

	message := messages[0]

	getItemConfig := ews.GetItemRequestConfig{
		ItemShape: &ews.ItemShape{
			BaseShape: ews.BaseShapeAllProperties,
			AdditionalProperties: &ews.AdditionalProperties{
				FieldURI: []ews.FieldURI{
					{
						FieldURI: "item:InternetMessageHeaders",
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
		return nil, errors.New("expected 1 message, got " + strconv.Itoa(len(messages)))
	}

	return &messages[0], nil
}
