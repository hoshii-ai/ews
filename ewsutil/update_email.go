package ewsutil

import (
	"github.com/hoshii-ai/ews"
	"github.com/hoshii-ai/ews/utils"
	"github.com/pkg/errors"
)

// UpdateEmailCategories updates the categories of an email by overwriting the existing categories.
func UpdateEmailCategories(c ews.Client, itemId *ews.ItemId, categories []string) (*ews.ItemId, error) {
	categories_ := ews.Categories{
		String: categories,
	}
	updateItemRequest := ews.UpdateItemRequest{
		MessageDisposition: ews.MessageDispositionSaveOnly,
		ConflictResolution: utils.Ptr(ews.ConflictResolutionAlwaysOverwrite),
		ItemChanges: ews.ItemChanges{
			ItemChange: []ews.ItemChange{
				{
					ItemId: *itemId,
					Updates: ews.Updates{
						SetItemField: []ews.SetItemField{
							{
								FieldURI: &ews.FieldURI{
									FieldURI: "item:Categories",
								},
								Message: &ews.Message{
									Categories: &categories_,
								},
							},
						},
					},
				},
			},
		},
	}

	updateItemResponse, err := ews.UpdateItem(c, &updateItemRequest)
	if err != nil {
		return nil, errors.Wrap(err, "failed to update item")
	}

	if ews.ResponseClass(updateItemResponse.ResponseMessages.UpdateItemResponseMessage.ResponseClass) != ews.ResponseClassSuccess {
		return nil, errors.New("failed to update item: " + updateItemResponse.ResponseMessages.UpdateItemResponseMessage.ResponseCode)
	}

	return itemId, nil
}
