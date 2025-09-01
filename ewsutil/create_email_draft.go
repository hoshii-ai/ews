package ewsutil

import (
	"github.com/hoshii-ai/ews"
	"github.com/hoshii-ai/ews/utils"
)

// SendEmail helper method to send Message
func CreateEmailDraft(c ews.Client, to []string, subject, body string, attachments ...ews.FileAttachment) (*ews.ItemId, error) {
	m := ews.Message{
		//ItemClass: utils.Ptr("IPM.Note"),
		Subject: utils.Ptr(subject),
		Body: &ews.Body{
			BodyType: "HTML",
			Body:     []byte(body),
		},
		// Sender: &ews.OneMailbox{
		// 	Mailbox: ews.Mailbox{
		// 		EmailAddress: c.GetUsername(),
		// 	},
		// },
		ToRecipients: &ews.XMailbox{
			Mailbox: make([]ews.Mailbox, len(to)),
		},
	}

	for i, addr := range to {
		m.ToRecipients.Mailbox[i].EmailAddress = addr
	}

	if len(attachments) > 0 {
		m.Attachments = &ews.Attachments{
			FileAttachment: attachments,
		}
	}

	return ews.CreateMessageItem(c, m, ews.CreateItemRequestConfig{
		MessageDisposition: ews.MessageDispositionSaveOnly,
		SavedItemFolderId:  &ews.SavedItemFolderId{DistinguishedFolderId: ews.DistinguishedFolderId{Id: "drafts"}},
	})
}
