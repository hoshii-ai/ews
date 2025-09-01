package ewsutil

import (
	"fmt"
	"os"
	"testing"

	"github.com/hoshii-ai/ews"
	"github.com/joho/godotenv"
)

var (
	url      string
	username string
	password string
)

func TestMain(m *testing.M) {
	_ = godotenv.Load("../.env")

	url = os.Getenv("EWS_URL")
	username = os.Getenv("EWS_USERNAME")
	password = os.Getenv("EWS_PASSWORD")

	os.Exit(m.Run())
}

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

	err := AddCategories(client, ews.Category{
		Name:  "test_hoshii_5",
		Color: ews.ColorBlue,
	})
	if err != nil {
		t.Fatalf("failed to add category: %v", err)
	}
}
