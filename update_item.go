package ews

import (
	"encoding/xml"
	"errors"
)

// UpdateItem
type ConflictResolution string

const (
	ConflictResolutionAutoResolve     ConflictResolution = "AutoResolve"
	ConflictResolutionAlwaysOverwrite ConflictResolution = "AlwaysOverwrite"
	ConflictResolutionNeverOverwrite  ConflictResolution = "NeverOverwrite"
)

type UpdateItemRequest struct {
	XMLName            xml.Name            `xml:"http://schemas.microsoft.com/exchange/services/2006/messages UpdateItem"`
	MessageDisposition MessageDisposition  `xml:"MessageDisposition,attr,omitempty"`
	ConflictResolution *ConflictResolution `xml:"ConflictResolution,attr,omitempty"`
	ItemChanges        ItemChanges         `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ItemChanges"`
}

type MessageDisposition string

const (
	MessageDispositionSaveOnly        MessageDisposition = "SaveOnly"
	MessageDispositionSendOnly        MessageDisposition = "SendOnly"
	MessageDispositionSendAndSaveCopy MessageDisposition = "SendAndSaveCopy"
)

type ItemChanges struct {
	ItemChange []ItemChange `xml:"http://schemas.microsoft.com/exchange/services/2006/types ItemChange"`
}

type ItemChange struct {
	ItemId  ItemId  `xml:"http://schemas.microsoft.com/exchange/services/2006/types ItemId"`
	Updates Updates `xml:"http://schemas.microsoft.com/exchange/services/2006/types Updates"`
}

type Updates struct {
	SetItemField      []SetItemField      `xml:"http://schemas.microsoft.com/exchange/services/2006/types SetItemField"`
	AppendToItemField []AppendToItemField `xml:"http://schemas.microsoft.com/exchange/services/2006/types AppendToItemField"`
	DeleteItemField   []DeleteItemField   `xml:"http://schemas.microsoft.com/exchange/services/2006/types DeleteItemField"`
}

type AppendToItemField struct {
	FieldURI FieldURI `xml:"http://schemas.microsoft.com/exchange/services/2006/types FieldURI"`
	Message  Message  `xml:"http://schemas.microsoft.com/exchange/services/2006/types Message"`
}

type SetItemField struct {
	FieldURI         *FieldURI         `xml:"http://schemas.microsoft.com/exchange/services/2006/types FieldURI,omitempty"`
	ExtendedFieldURI *ExtendedFieldURI `xml:"http://schemas.microsoft.com/exchange/services/2006/types ExtendedFieldURI,omitempty"`
	Message          *Message          `xml:"http://schemas.microsoft.com/exchange/services/2006/types Message,omitempty"`
}

type DeleteItemField struct {
	FieldURI         *FieldURI         `xml:"http://schemas.microsoft.com/exchange/services/2006/types FieldURI,omitempty"`
	ExtendedFieldURI *ExtendedFieldURI `xml:"http://schemas.microsoft.com/exchange/services/2006/types ExtendedFieldURI,omitempty"`
}

// UpdateItemResponse (minimal for now, can be expanded later)

type UpdateItemResponseEnvelope struct {
	XMLName xml.Name               `xml:"Envelope"`
	Header  ServerVersionInfo      `xml:"Header"`
	Body    UpdateItemResponseBody `xml:"Body"`
}

type UpdateItemResponseBody struct {
	UpdateItemResponse UpdateItemResponse `xml:"http://schemas.microsoft.com/exchange/services/2006/messages UpdateItemResponse"`
}

type UpdateItemResponse struct {
	ResponseMessages UpdateItemResponseMessages `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
}

type UpdateItemResponseMessages struct {
	UpdateItemResponseMessage UpdateItemResponseMessage `xml:"http://schemas.microsoft.com/exchange/services/2006/messages UpdateItemResponseMessage"`
}

type UpdateItemResponseMessage struct {
	ResponseClass ResponseClass `xml:"ResponseClass,attr"`
	ResponseCode  string        `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseCode"`
}

// UpdateItem takes an UpdateItem request and returns an UpdateItemResponse.
// https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/updateitem-operation
func UpdateItem(c Client, r *UpdateItemRequest) (*UpdateItemResponse, error) {
	xmlBytes, err := xml.MarshalIndent(r, "", "  ")
	if err != nil {
		return nil, err
	}

	bb, err := c.SendAndReceive(xmlBytes)
	if err != nil {
		return nil, err
	}

	var soapResp UpdateItemResponseEnvelope
	err = xml.Unmarshal(bb, &soapResp)
	if err != nil {
		return nil, err
	}

	if ResponseClass(soapResp.Body.UpdateItemResponse.ResponseMessages.UpdateItemResponseMessage.ResponseClass) == ResponseClassError {
		return nil, errors.New(soapResp.Body.UpdateItemResponse.ResponseMessages.UpdateItemResponseMessage.ResponseCode)
	}

	return &soapResp.Body.UpdateItemResponse, nil
}
