package ewsutil

import (
	"testing"

	"github.com/hoshii-ai/ews"
)

func Test_UpdateEmailCategories(t *testing.T) {
	c := ews.NewClient(url, username, password, &ews.Config{
		Dump:    true,
		NTLM:    true,
		SkipTLS: false,
	})

	categories, err := GetInboxCategories(c)
	if err != nil {
		t.Fatalf("failed to get inbox categories: %v", err)
	}

	email, err := GetMessageByInternetMessageId(c, "<CAKMBqB1yyODh+Bjpv9bL7attSoKxPo-a03t_N5QKKt_+eBPh1w@mail.gmail.com>")
	if err != nil {
		t.Fatalf("failed to get email: %v", err)
	}

	missingCategory := ""
	for _, category := range categories.Categories {
		for _, emailCategory := range email.Categories.String {
			if category.Name != emailCategory {
				missingCategory = category.Name
				break
			}
		}
	}

	if missingCategory == "" {
		t.Fatalf("no missing category found")
	}

	_, err = UpdateEmailCategories(c, email.ItemId, append(email.Categories.String, missingCategory))
	if err != nil {
		t.Fatalf("failed to update email categories: %v", err)
	}
	t.Logf("updated email: %+v", *email.Subject)
	t.Logf("updated email categories: %s", missingCategory)
}
