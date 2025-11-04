package ewsutil

import (
	"github.com/hoshii-ai/ews"
	"github.com/hoshii-ai/ews/utils"
	"github.com/pkg/errors"
)

// SendEmail helper method to send Message
func SendEmail(c ews.Client, to []string, subject, body string, attachments ...ews.FileAttachment) (*ews.ItemId, error) {
	m := ews.Message{
		Subject: utils.Ptr(subject),
		Body: &ews.Body{
			BodyType: "HTML",
			Body:     []byte(body),
		},
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

	itemId, err := sendEmailWithSaveThenSend(c, m)
	if err != nil {
		return nil, errors.Wrap(err, "failed to send email")
	}

	return itemId, nil
}

// This function is deprecated: some EWS servers do not return the item id after successfully creating the message item
// func sendEmailWithAtomicSaveAndSend(c ews.Client, m ews.Message) (*ews.ItemId, error) {
// 	return ews.CreateMessageItem(c, m, ews.CreateItemRequestConfig{
// 		MessageDisposition: ews.MessageDispositionSendAndSaveCopy,
// 		SavedItemFolderId:  &ews.SavedItemFolderId{DistinguishedFolderId: ews.DistinguishedFolderId{Id: "sentitems"}},
// 	})
// }

func sendEmailWithSaveThenSend(c ews.Client, m ews.Message) (*ews.ItemId, error) {
	// Save the email draft first
	itemId, err := ews.CreateMessageItem(c, m, ews.CreateItemRequestConfig{
		MessageDisposition: ews.MessageDispositionSendOnly,
		SavedItemFolderId:  &ews.SavedItemFolderId{DistinguishedFolderId: ews.DistinguishedFolderId{Id: "drafts"}},
	})
	if err != nil {
		return nil, errors.Wrap(err, "failed to create message item")
	}

	// Send the email draft
	if err := SendEmailWithItemId(c, itemId); err != nil {
		return nil, errors.Wrap(err, "failed to send email")
	}

	return itemId, nil
}

func SendEmailWithItemId(c ews.Client, itemId *ews.ItemId) error {
	_, err := ews.SendItem(c, *itemId, true)
	if err != nil {
		return errors.Wrap(err, "failed to send email")
	}

	return nil
}
