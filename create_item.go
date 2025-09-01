package ews

import (
	"encoding/xml"
	"strconv"
	"time"

	"github.com/pkg/errors"
)

type CreateItemRequest struct {
	XMLName                struct{}           `xml:"http://schemas.microsoft.com/exchange/services/2006/messages CreateItem"`
	MessageDisposition     MessageDisposition `xml:"MessageDisposition,attr"`
	SendMeetingInvitations *string            `xml:"SendMeetingInvitations,attr,omitempty"`
	SavedItemFolderId      *SavedItemFolderId `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SavedItemFolderId,omitempty"`
	Items                  Items              `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Items"`
}

func NewCreateItemRequest(item any, config CreateItemRequestConfig) (*CreateItemRequest, error) {
	switch item := item.(type) {
	case Message:
		return &CreateItemRequest{
			MessageDisposition: config.MessageDisposition,
			SavedItemFolderId:  config.SavedItemFolderId,
			Items: Items{
				Message: []Message{item},
			},
		}, nil
	case CalendarItem:
		return &CreateItemRequest{
			MessageDisposition: config.MessageDisposition,
			SavedItemFolderId:  config.SavedItemFolderId,
			Items: Items{
				CalendarItem: []CalendarItem{item},
			},
		}, nil
	}

	return nil, errors.New("invalid item type")
}

type Items struct {
	Message      []Message      `xml:"http://schemas.microsoft.com/exchange/services/2006/types Message"`
	CalendarItem []CalendarItem `xml:"http://schemas.microsoft.com/exchange/services/2006/types CalendarItem"`
}

type SavedItemFolderId struct {
	DistinguishedFolderId DistinguishedFolderId `xml:"http://schemas.microsoft.com/exchange/services/2006/types DistinguishedFolderId"`
}

type Message struct {
	ItemId                       *ItemId                 `xml:"http://schemas.microsoft.com/exchange/services/2006/types ItemId,omitempty"`
	ParentFolderId               *ParentFolderId         `xml:"http://schemas.microsoft.com/exchange/services/2006/types ParentFolderId,omitempty"`
	Categories                   *Categories             `xml:"http://schemas.microsoft.com/exchange/services/2006/types Categories,omitempty"`
	ItemClass                    *string                 `xml:"http://schemas.microsoft.com/exchange/services/2006/types ItemClass,omitempty"`
	Subject                      *string                 `xml:"http://schemas.microsoft.com/exchange/services/2006/types Subject,omitempty"`
	Sensitivity                  *string                 `xml:"http://schemas.microsoft.com/exchange/services/2006/types Sensitivity,omitempty"` // TODO: enum
	Body                         *Body                   `xml:"http://schemas.microsoft.com/exchange/services/2006/types Body,omitempty"`
	DateTimeReceived             *time.Time              `xml:"http://schemas.microsoft.com/exchange/services/2006/types DateTimeReceived,omitempty"`
	Size                         *int                    `xml:"http://schemas.microsoft.com/exchange/services/2006/types Size,omitempty"`
	Importance                   *string                 `xml:"http://schemas.microsoft.com/exchange/services/2006/types Importance,omitempty"`
	IsSubmitted                  *bool                   `xml:"http://schemas.microsoft.com/exchange/services/2006/types IsSubmitted,omitempty"`
	IsDraft                      *bool                   `xml:"http://schemas.microsoft.com/exchange/services/2006/types IsDraft,omitempty"`
	IsFromMe                     *bool                   `xml:"http://schemas.microsoft.com/exchange/services/2006/types IsFromMe,omitempty"`
	IsResend                     *bool                   `xml:"http://schemas.microsoft.com/exchange/services/2006/types IsResend,omitempty"`
	IsUnmodified                 *bool                   `xml:"http://schemas.microsoft.com/exchange/services/2006/types IsUnmodified,omitempty"`
	InternetMessageHeaders       *InternetMessageHeaders `xml:"http://schemas.microsoft.com/exchange/services/2006/types InternetMessageHeaders,omitempty"`
	DateTimeSent                 *time.Time              `xml:"http://schemas.microsoft.com/exchange/services/2006/types DateTimeSent,omitempty"`
	DateTimeCreated              *time.Time              `xml:"http://schemas.microsoft.com/exchange/services/2006/types DateTimeCreated,omitempty"`
	ResponseObjects              *ResponseObjects        `xml:"http://schemas.microsoft.com/exchange/services/2006/types ResponseObjects,omitempty"`
	DisplayCc                    *string                 `xml:"http://schemas.microsoft.com/exchange/services/2006/types DisplayCc,omitempty"`
	DisplayTo                    *string                 `xml:"http://schemas.microsoft.com/exchange/services/2006/types DisplayTo,omitempty"`
	HasAttachments               *bool                   `xml:"http://schemas.microsoft.com/exchange/services/2006/types HasAttachments,omitempty"`
	Attachments                  *Attachments            `xml:"http://schemas.microsoft.com/exchange/services/2006/types Attachments,omitempty"`
	Culture                      *string                 `xml:"http://schemas.microsoft.com/exchange/services/2006/types Culture,omitempty"`
	EffectiveRights              *EffectiveRights        `xml:"http://schemas.microsoft.com/exchange/services/2006/types EffectiveRights,omitempty"`
	LastModifiedName             *string                 `xml:"http://schemas.microsoft.com/exchange/services/2006/types LastModifiedName,omitempty"`
	LastModifiedTime             *time.Time              `xml:"http://schemas.microsoft.com/exchange/services/2006/types LastModifiedTime,omitempty"`
	IsAssociated                 *bool                   `xml:"http://schemas.microsoft.com/exchange/services/2006/types IsAssociated,omitempty"`
	WebClientReadFormQueryString *string                 `xml:"http://schemas.microsoft.com/exchange/services/2006/types WebClientReadFormQueryString,omitempty"`
	ConversationId               *ConversationId         `xml:"http://schemas.microsoft.com/exchange/services/2006/types ConversationId,omitempty"`
	Flag                         *Flag                   `xml:"http://schemas.microsoft.com/exchange/services/2006/types Flag,omitempty"`
	InstanceKey                  *string                 `xml:"http://schemas.microsoft.com/exchange/services/2006/types InstanceKey,omitempty"`
	EntityExtractionResult       *struct{}               `xml:"http://schemas.microsoft.com/exchange/services/2006/types EntityExtractionResult,omitempty"`
	Sender                       *OneMailbox             `xml:"http://schemas.microsoft.com/exchange/services/2006/types Sender,omitempty"`
	ToRecipients                 *XMailbox               `xml:"http://schemas.microsoft.com/exchange/services/2006/types ToRecipients,omitempty"`
	IsReadReceiptRequested       *bool                   `xml:"http://schemas.microsoft.com/exchange/services/2006/types IsReadReceiptRequested,omitempty"`
	ConversationIndex            *string                 `xml:"http://schemas.microsoft.com/exchange/services/2006/types ConversationIndex,omitempty"`
	ConversationTopic            *string                 `xml:"http://schemas.microsoft.com/exchange/services/2006/types ConversationTopic,omitempty"`
	From                         *OneMailbox             `xml:"http://schemas.microsoft.com/exchange/services/2006/types From,omitempty"`
	InternetMessageId            *string                 `xml:"http://schemas.microsoft.com/exchange/services/2006/types InternetMessageId,omitempty"`
	IsRead                       *bool                   `xml:"http://schemas.microsoft.com/exchange/services/2006/types IsRead,omitempty"`
	ReceivedBy                   *OneMailbox             `xml:"http://schemas.microsoft.com/exchange/services/2006/types ReceivedBy,omitempty"`
	ReceivedRepresenting         *OneMailbox             `xml:"http://schemas.microsoft.com/exchange/services/2006/types ReceivedRepresenting,omitempty"`

	ExtendedProperties []ExtendedProperty `xml:"http://schemas.microsoft.com/exchange/services/2006/types ExtendedProperty,omitempty"`
}

func (m *Message) GetHeaders() (map[string]string, error) {
	headers := make(map[string]string)
	if m.InternetMessageHeaders == nil {
		return headers, errors.New("no internet message headers")
	}
	for _, header := range m.InternetMessageHeaders.InternetMessageHeader {
		headers[header.HeaderName] = header.Value
	}
	return headers, nil
}

type ExtendedProperty struct {
	ExtendedFieldURI *ExtendedFieldURI `xml:"http://schemas.microsoft.com/exchange/services/2006/types ExtendedFieldURI,omitempty"`
	FieldURI         *FieldURI         `xml:"http://schemas.microsoft.com/exchange/services/2006/types FieldURI,omitempty"`
	Value            *string           `xml:"http://schemas.microsoft.com/exchange/services/2006/types Value,omitempty"`
}

type Categories struct {
	String []string `xml:"http://schemas.microsoft.com/exchange/services/2006/types String"`
}

type CalendarItem struct {
	Subject                    string      `xml:"t:Subject"`
	Body                       Body        `xml:"t:Body"`
	ReminderIsSet              bool        `xml:"t:ReminderIsSet"`
	ReminderMinutesBeforeStart int         `xml:"t:ReminderMinutesBeforeStart"`
	Start                      time.Time   `xml:"t:Start"`
	End                        time.Time   `xml:"t:End"`
	IsAllDayEvent              bool        `xml:"t:IsAllDayEvent"`
	LegacyFreeBusyStatus       string      `xml:"t:LegacyFreeBusyStatus"`
	Location                   string      `xml:"t:Location"`
	RequiredAttendees          []Attendees `xml:"t:RequiredAttendees"`
	OptionalAttendees          []Attendees `xml:"t:OptionalAttendees"`
	Resources                  []Attendees `xml:"t:Resources"`
}

type Body struct {
	BodyType string `xml:"BodyType,attr"`
	Body     []byte `xml:",chardata"`
}

type OneMailbox struct {
	Mailbox Mailbox `xml:"http://schemas.microsoft.com/exchange/services/2006/types Mailbox"`
}

type XMailbox struct {
	Mailbox []Mailbox `xml:"http://schemas.microsoft.com/exchange/services/2006/types Mailbox"`
}

type Mailbox struct {
	EmailAddress string `xml:"http://schemas.microsoft.com/exchange/services/2006/types EmailAddress"`
}

type Attendee struct {
	Mailbox Mailbox `xml:"http://schemas.microsoft.com/exchange/services/2006/types Mailbox"`
}

type Attendees struct {
	Attendee []Attendee `xml:"http://schemas.microsoft.com/exchange/services/2006/types Attendee"`
}

type createItemResponseBodyEnvelope struct {
	XMLName struct{}               `xml:"Envelope"`
	Body    createItemResponseBody `xml:"Body"`
}
type createItemResponseBody struct {
	CreateItemResponse CreateItemResponse `xml:"http://schemas.microsoft.com/exchange/services/2006/messages CreateItemResponse"`
}

type CreateItemResponse struct {
	ResponseMessages ResponseMessages `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
}

type ResponseMessages struct {
	CreateItemResponseMessage Response `xml:"http://schemas.microsoft.com/exchange/services/2006/messages CreateItemResponseMessage"`
}

type CreateItemRequestConfig struct {
	MessageDisposition MessageDisposition
	SavedItemFolderId  *SavedItemFolderId
}

// CreateMessageItem
// https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/createitem-operation-email-message
func CreateMessageItem(c Client, message Message, config CreateItemRequestConfig) (*ItemId, error) {
	createItemRequest, err := NewCreateItemRequest(message, config)
	if err != nil {
		return nil, errors.Wrap(err, "failed to create create item request")
	}

	xmlBytes, err := xml.MarshalIndent(createItemRequest, "", "  ")
	if err != nil {
		return nil, err
	}

	bb, err := c.SendAndReceive(xmlBytes)
	if err != nil {
		return nil, err
	}

	var soapResp createItemResponseBodyEnvelope
	if err := xml.Unmarshal(bb, &soapResp); err != nil {
		return nil, err
	}

	resp := soapResp.Body.CreateItemResponse.ResponseMessages.CreateItemResponseMessage
	if resp.ResponseClass == ResponseClassError {
		return nil, errors.New(resp.MessageText)
	}

	messages := resp.Items.Message
	if len(messages) != 1 {
		return nil, errors.New("expected 1 message, got " + strconv.Itoa(len(messages)))
	}

	return messages[0].ItemId, nil
}

// CreateCalendarItem
// https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/createitem-operation-calendar-item
func CreateCalendarItem(c Client, calendarItem CalendarItem) error {
	createItemRequest, err := NewCreateItemRequest(calendarItem, CreateItemRequestConfig{
		MessageDisposition: MessageDispositionSendAndSaveCopy,
		SavedItemFolderId:  &SavedItemFolderId{DistinguishedFolderId{Id: "calendar"}},
	})
	if err != nil {
		return errors.Wrap(err, "failed to create create item request")
	}

	xmlBytes, err := xml.MarshalIndent(createItemRequest, "", "  ")
	if err != nil {
		return err
	}

	bb, err := c.SendAndReceive(xmlBytes)
	if err != nil {
		return err
	}

	var soapResp createItemResponseBodyEnvelope
	if err := xml.Unmarshal(bb, &soapResp); err != nil {
		return err
	}

	resp := soapResp.Body.CreateItemResponse.ResponseMessages.CreateItemResponseMessage
	if resp.ResponseClass == ResponseClassError {
		return errors.New(resp.MessageText)
	}

	return nil
}
