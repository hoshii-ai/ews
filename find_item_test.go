package ews

import (
	"fmt"
	"testing"
)

func TestFindItem(t *testing.T) {
	t.Run("should find item", func(t *testing.T) {
		client := NewClient(url, username, password, &Config{
			Dump:    true,
			NTLM:    true,
			SkipTLS: false,
		})

		item, err := FindItem(client,
			"inbox",
			FindItemRequestConfig{},
		)
		if err != nil {
			t.Fatalf("failed to get item: %v", err)
		}

		fmt.Println(item)
		// fmt.Println(item.RootFolder.TotalItemsInView)
	})
}
