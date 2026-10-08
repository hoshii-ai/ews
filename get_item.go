package ews

import (
	"context"
	"encoding/base64"
	"encoding/xml"
	"errors"
	"strings"
)

type GetItemRequest struct {
	XMLName   xml.Name  `xml:"http://schemas.microsoft.com/exchange/services/2006/messages GetItem"`
	ItemShape ItemShape `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ItemShape"`
	ItemIds   ItemIds   `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ItemIds"`
}

type GetItemRequestConfig struct {
	ItemShape *ItemShape
}

func NewGetItemRequest(itemId ItemId, config GetItemRequestConfig) *GetItemRequest {
	itemShape := ItemShape{
		BaseShape: BaseShapeAllProperties,
	}
	if config.ItemShape != nil {
		itemShape = *config.ItemShape
	}

	return &GetItemRequest{
		ItemShape: itemShape,
		ItemIds: ItemIds{
			ItemId: []ItemId{itemId},
		},
	}
}

type ItemIds struct {
	ItemId []ItemId `xml:"http://schemas.microsoft.com/exchange/services/2006/types ItemId"`
}

type GetItemResponseEnvelope struct {
	XMLName xml.Name            `xml:"Envelope"`
	Header  ServerVersionInfo   `xml:"Header"`
	Body    GetItemResponseBody `xml:"Body"`
}

type GetItemResponseBody struct {
	GetItemResponse GetItemResponse `xml:"http://schemas.microsoft.com/exchange/services/2006/messages GetItemResponse"`
}

type GetItemResponse struct {
	ResponseMessages GetItemResponseMessages `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
}

type GetItemResponseMessages struct {
	GetItemResponseMessage GetItemResponseMessage `xml:"http://schemas.microsoft.com/exchange/services/2006/messages GetItemResponseMessage"`
}

type GetItemResponseMessage struct {
	ResponseClass ResponseClass `xml:"ResponseClass,attr"`
	ResponseCode  string        `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseCode"`
	MessageText   string        `xml:"http://schemas.microsoft.com/exchange/services/2006/messages MessageText"`
	MessageXml    MessageXml    `xml:"http://schemas.microsoft.com/exchange/services/2006/messages MessageXml"`
	Items         Items         `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Items"`
}

// MimeContent is the MIME stream of an item, returned when ItemShape.IncludeMimeContent is set.
type MimeContent struct {
	CharacterSet string `xml:"CharacterSet,attr"`
	// Value is base64-encoded MIME.
	Value string `xml:",chardata"`
}

// Decode returns the raw MIME bytes.
func (m *MimeContent) Decode() ([]byte, error) {
	if m == nil {
		return nil, errors.New("ews: no MimeContent")
	}
	return base64.StdEncoding.DecodeString(strings.Join(strings.Fields(m.Value), ""))
}

// NewGetItemMimeShape returns an ItemShape that fetches the MIME content only (IdOnly + IncludeMimeContent).
func NewGetItemMimeShape() *ItemShape {
	return &ItemShape{BaseShape: BaseShapeIdOnly, IncludeMimeContent: true}
}

type GetItemBody struct {
	BodyType    string `xml:"BodyType,attr"`
	IsTruncated bool   `xml:"IsTruncated,attr"`
	Body        string `xml:",chardata"`
}

type InternetMessageHeaders struct {
	InternetMessageHeader []InternetMessageHeader `xml:"http://schemas.microsoft.com/exchange/services/2006/types InternetMessageHeader"`
}

type InternetMessageHeader struct {
	HeaderName string `xml:"HeaderName,attr"`
	Value      string `xml:",chardata"`
}

type ResponseObjects struct {
	ReplyToItem    struct{} `xml:"ReplyToItem"`
	ReplyAllToItem struct{} `xml:"ReplyAllToItem"`
	ForwardItem    struct{} `xml:"ForwardItem"`
}

type EffectiveRights struct {
	CreateAssociated bool `xml:"CreateAssociated"`
	CreateContents   bool `xml:"CreateContents"`
	CreateHierarchy  bool `xml:"CreateHierarchy"`
	Delete           bool `xml:"Delete"`
	Modify           bool `xml:"Modify"`
	Read             bool `xml:"Read"`
	ViewPrivateItems bool `xml:"ViewPrivateItems"`
}

type Flag struct {
	FlagStatus string `xml:"FlagStatus"`
}

type ConversationId struct {
	Id string `xml:"Id,attr"`
}

// GetItem takes a GetItemRequest and returns a GetItemResponse
// https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/getitem-operation
func GetItem(c Client, itemId ItemId, config GetItemRequestConfig) (*GetItemResponse, error) {
	r := NewGetItemRequest(itemId, config)
	xmlBytes, err := xml.MarshalIndent(r, "", "  ")
	if err != nil {
		return nil, err
	}

	bb, err := c.SendAndReceive(xmlBytes)
	if err != nil {
		return nil, err
	}

	var soapResp GetItemResponseEnvelope
	err = xml.Unmarshal(bb, &soapResp)
	if err != nil {
		return nil, err
	}

	if ResponseClass(soapResp.Body.GetItemResponse.ResponseMessages.GetItemResponseMessage.ResponseClass) == ResponseClassError {
		return nil, errors.New(soapResp.Body.GetItemResponse.ResponseMessages.GetItemResponseMessage.ResponseCode)
	}

	return &soapResp.Body.GetItemResponse, nil
}

// GetItemContext is GetItem with a context and per-request options (e.g.
// WithAnchorMailbox). ResponseClass=Error is returned as *ResponseError.
func GetItemContext(ctx context.Context, c ContextClient, itemId ItemId, config GetItemRequestConfig, opts ...RequestOption) (*GetItemResponse, error) {
	var env GetItemResponseEnvelope
	if err := call(ctx, c, NewGetItemRequest(itemId, config), &env, opts); err != nil {
		return nil, err
	}
	m := env.Body.GetItemResponse.ResponseMessages.GetItemResponseMessage
	if err := (responseMessage{ResponseClass: m.ResponseClass, ResponseCode: m.ResponseCode, MessageText: m.MessageText, MessageXml: m.MessageXml}).err(); err != nil {
		return nil, err
	}
	return &env.Body.GetItemResponse, nil
}
