package ews

import (
	"encoding/xml"
	"errors"
)

type AttachmentShape struct {
	IncludeMimeContent   bool       `xml:"m:IncludeMimeContent,omitempty"`
	BodyType             string     `xml:"m:BodyType,omitempty"`
	FilterHtmlContent    bool       `xml:"m:FilterHtmlContent,omitempty"`
	AdditionalProperties []FieldURI `xml:"m:AdditionalProperties>t:FieldURI,omitempty"`
}

type AttachmentIds struct {
	AttachmentId []AttachmentId `xml:"t:AttachmentId"`
}

type GetAttachmentRequest struct {
	XMLName         xml.Name        `xml:"m:GetAttachment"`
	AttachmentShape AttachmentShape `xml:"m:AttachmentShape"`
	AttachmentIds   AttachmentIds   `xml:"m:AttachmentIds"`
}

type GetAttachmentResponseEnvelope struct {
	XMLName xml.Name `xml:"soap:Envelope"`
	Header  struct {
		ServerVersionInfo ServerVersionInfo `xml:"h:ServerVersionInfo"` // Uncommented and added h: prefix
	} `xml:"soap:Header"`
	Body GetAttachmentResponseBody `xml:"soap:Body"`
}

type GetAttachmentResponseBody struct {
	GetAttachmentResponse GetAttachmentResponse `xml:"m:GetAttachmentResponse"`
}

type GetAttachmentResponse struct {
	ResponseMessages GetAttachmentResponseMessages `xml:"m:ResponseMessages"`
}

type GetAttachmentResponseMessages struct {
	GetAttachmentResponseMessage GetAttachmentResponseMessage `xml:"m:GetAttachmentResponseMessage"`
}

type GetAttachmentResponseMessage struct {
	ResponseClass string      `xml:"ResponseClass,attr"`
	ResponseCode  string      `xml:"ResponseCode"`
	Attachments   Attachments `xml:"Attachments"`
}

type Attachments struct {
	ItemAttachment []ItemAttachment `xml:"http://schemas.microsoft.com/exchange/services/2006/types ItemAttachment,omitempty"`
	FileAttachment []FileAttachment `xml:"http://schemas.microsoft.com/exchange/services/2006/types FileAttachment,omitempty"`
}

type ItemAttachment struct {
	// ItemAttachment typically contains a full Item or Message structure, but based on your XML, it's empty for now.
	// We'll keep it empty and add fields if needed later.
}

type FileAttachment struct {
	AttachmentId     *AttachmentId `xml:"http://schemas.microsoft.com/exchange/services/2006/types AttachmentId,omitempty"`
	Name             string        `xml:"http://schemas.microsoft.com/exchange/services/2006/types Name"`
	ContentType      string        `xml:"http://schemas.microsoft.com/exchange/services/2006/types ContentType,omitempty"`
	ContentId        string        `xml:"http://schemas.microsoft.com/exchange/services/2006/types ContentId,omitempty"`
	ContentLocation  string        `xml:"http://schemas.microsoft.com/exchange/services/2006/types ContentLocation,omitempty"`
	Size             int64         `xml:"http://schemas.microsoft.com/exchange/services/2006/types Size,omitempty"`
	LastModifiedTime string        `xml:"http://schemas.microsoft.com/exchange/services/2006/types LastModifiedTime,omitempty"`
	IsInline         *bool         `xml:"http://schemas.microsoft.com/exchange/services/2006/types IsInline,omitempty"`
	IsContactPhoto   *bool         `xml:"http://schemas.microsoft.com/exchange/services/2006/types IsContactPhoto,omitempty"`
	Content          string        `xml:"http://schemas.microsoft.com/exchange/services/2006/types Content"` // Base64 encoded content
}

// GetAttachment takes a GetAttachmentRequest and returns a GetAttachmentResponse
// https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/getattachment-operation
func GetAttachment(c Client, r *GetAttachmentRequest) (*GetAttachmentResponse, error) {
	xmlBytes, err := xml.MarshalIndent(r, "", "  ")
	if err != nil {
		return nil, err
	}

	bb, err := c.SendAndReceive(xmlBytes)
	if err != nil {
		return nil, err
	}

	var soapResp GetAttachmentResponseEnvelope
	err = xml.Unmarshal(bb, &soapResp)
	if err != nil {
		return nil, err
	}

	if ResponseClass(soapResp.Body.GetAttachmentResponse.ResponseMessages.GetAttachmentResponseMessage.ResponseClass) == ResponseClassError {
		return nil, errors.New(soapResp.Body.GetAttachmentResponse.ResponseMessages.GetAttachmentResponseMessage.ResponseCode)
	}

	return &soapResp.Body.GetAttachmentResponse, nil
}
