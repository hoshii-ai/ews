package ews

import (
	"context"
	"encoding/xml"
	"strconv"
)

type SendItemRequest struct {
	XMLName          struct{} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SendItem"`
	SaveItemToFolder string   `xml:"SaveItemToFolder,attr"` // true or false
	ItemIds          ItemIds  `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ItemIds"`
}

// --- Response ---

// --- SOAP envelope for unmarshaling ---

type sendItemResponseEnvelope struct {
	XMLName xml.Name             `xml:"Envelope"`
	Header  ServerVersionInfo    `xml:"Header"`
	Body    sendItemResponseBody `xml:"Body"`
}

type sendItemResponseBody struct {
	SendItemResponse SendItemResponse `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SendItemResponse"`
}
type SendItemResponse struct {
	ResponseMessages SendItemResponseMessages `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
}

type SendItemResponseMessages struct {
	SendItemResponseMessage SendItemResponseMessage `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SendItemResponseMessage"`
}

type SendItemResponseMessage struct {
	ResponseClass ResponseClass `xml:"ResponseClass,attr"`
	ResponseCode  string        `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseCode"`
	MessageText   string        `xml:"http://schemas.microsoft.com/exchange/services/2006/messages MessageText"`
	MessageXml    MessageXml    `xml:"http://schemas.microsoft.com/exchange/services/2006/messages MessageXml"`
}

func (m SendItemResponseMessage) err() error {
	return responseMessage{ResponseClass: m.ResponseClass, ResponseCode: m.ResponseCode, MessageText: m.MessageText, MessageXml: m.MessageXml}.err()
}

// --- Example function to send request ---
func SendItem(c Client, itemId ItemId, saveItemToFolder bool) (*SendItemResponse, error) {
	saveItemToFolderStr := "true"
	if !saveItemToFolder {
		saveItemToFolderStr = "false"
	}
	req := SendItemRequest{
		SaveItemToFolder: saveItemToFolderStr,
		ItemIds:          ItemIds{ItemId: []ItemId{itemId}},
	}
	xmlBytes, err := xml.MarshalIndent(req, "", "  ")
	if err != nil {
		return nil, err
	}

	bb, err := c.SendAndReceive(xmlBytes)
	if err != nil {
		return nil, err
	}

	var soapResp sendItemResponseEnvelope

	err = xml.Unmarshal(bb, &soapResp)
	if err != nil {
		return nil, err
	}

	if err := soapResp.Body.SendItemResponse.ResponseMessages.SendItemResponseMessage.err(); err != nil {
		return nil, err
	}

	return &soapResp.Body.SendItemResponse, nil
}

// SendItemContext is SendItem with a context and per-request options.
// ResponseClass=Error is returned as *ResponseError.
func SendItemContext(ctx context.Context, c ContextClient, itemId ItemId, saveItemToFolder bool, opts ...RequestOption) error {
	req := SendItemRequest{
		SaveItemToFolder: strconv.FormatBool(saveItemToFolder),
		ItemIds:          ItemIds{ItemId: []ItemId{itemId}},
	}
	var env sendItemResponseEnvelope
	if err := call(ctx, c, req, &env, opts); err != nil {
		return err
	}
	return env.Body.SendItemResponse.ResponseMessages.SendItemResponseMessage.err()
}
