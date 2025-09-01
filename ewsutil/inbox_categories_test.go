package ewsutil

import (
	"fmt"
	"os"
	"testing"

	"github.com/hoshii-ai/ews"
)

var (
	url      = os.Getenv("EWS_URL")
	username = os.Getenv("EWS_USERNAME")
	password = os.Getenv("EWS_PASSWORD")
)

func Test_GetInboxCategories(t *testing.T) {
	client := ews.NewClient(url, username, password, &ews.Config{
		Dump:    true,
		NTLM:    true,
		SkipTLS: false,
	})

	list, err := GetInboxCategories(client)
	if err != nil {
		t.Fatalf("failed to get inbox categories: %v", err)
	}

	for _, category := range list.Categories {
		fmt.Println("--------------------------------")
		fmt.Println(category.Name)
		fmt.Println(category.Color)
		fmt.Println(category.KeyboardShortcut)
		fmt.Println(category.UsageCount)
		fmt.Println(category.LastTimeUsed)
		fmt.Println(category.LastSessionUsed)
		fmt.Println(category.GUID)
		fmt.Println(category.RenameOnFirstUse)
		fmt.Println("--------------------------------")
	}
}

func Test_AddCategory(t *testing.T) {
	client := ews.NewClient(url, username, password, &ews.Config{
		Dump:    true,
		NTLM:    true,
		SkipTLS: false,
	})

	categories, err := GetInboxCategories(client)
	if err != nil {
		t.Fatalf("failed to get inbox categories: %v", err)
	}

	err = categories.AddCategory("this_comes_from_hoshii_3", ews.ColorPurple)
	if err != nil {
		t.Fatalf("failed to add category: %v", err)
	}

	xmlData, err := categories.CategoryListToBase64()
	if err != nil {
		t.Fatalf("failed to convert category list to base64: %v", err)
	}

	fmt.Println(xmlData)

	updateItemRequest := &ews.UpdateItemRequest{
		MessageDisposition: ews.MessageDispositionSaveOnly,
		ItemChanges: ews.ItemChanges{
			ItemChange: []ews.ItemChange{
				{
					ItemId: categories.ItemId,
					Updates: ews.Updates{
						SetItemField: []ews.SetItemField{
							{
								ExtendedFieldURI: &ews.ExtendedFieldURI{
									PropertyTag:  ews.PropertyTagCategories,
									PropertyType: ews.PropertyTypeBinary,
								},
								Message: &ews.Message{
									ExtendedProperties: []ews.ExtendedProperty{
										{
											Value: &xmlData,
										},
									},
								},
							},
						},
					},
				},
			},
		},
	}

	updateItemResponse, err := ews.UpdateItem(client, updateItemRequest)
	if err != nil {
		t.Fatalf("failed to update item: %v", err)
	}

	fmt.Println(updateItemResponse)
}
