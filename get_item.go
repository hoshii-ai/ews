package ews

import (
	"encoding/xml"
	"errors"
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
	Items         Items         `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Items"`
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
