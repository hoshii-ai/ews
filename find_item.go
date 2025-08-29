package ews

import (
	"encoding/xml"

	"github.com/hoshii-ai/ews/utils"
)

// --- Request ---

type FindItemRequest struct {
	XMLName         struct{}      `xml:"http://schemas.microsoft.com/exchange/services/2006/messages FindItem"`
	Traversal       string        `xml:"Traversal,attr,omitempty"`
	ItemShape       ItemShape     `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ItemShape"`
	Restriction     *Restriction  `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Restriction,omitempty"`
	ParentFolderIds ParentFolders `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ParentFolderIds"`
	QueryString     *string       `xml:"http://schemas.microsoft.com/exchange/services/2006/messages QueryString,omitempty"`
}

type Restriction struct {
	IsEqualTo *IsEqualTo `xml:"http://schemas.microsoft.com/exchange/services/2006/types IsEqualTo,omitempty"`
}

type IsEqualTo struct {
	FieldURI           *FieldURI           `xml:"http://schemas.microsoft.com/exchange/services/2006/types FieldURI,omitempty"`
	ExtendedFieldURI   *ExtendedFieldURI   `xml:"http://schemas.microsoft.com/exchange/services/2006/types ExtendedFieldURI,omitempty"`
	FieldURIOrConstant *FieldURIOrConstant `xml:"http://schemas.microsoft.com/exchange/services/2006/types FieldURIOrConstant"`
}

type FieldURIOrConstant struct {
	Constant *Constant `xml:"http://schemas.microsoft.com/exchange/services/2006/types Constant,omitempty"`
}

type Constant struct {
	Value string `xml:"Value,attr"`
}
type ParentFolders struct {
	FolderIds []DistinguishedFolderId `xml:"http://schemas.microsoft.com/exchange/services/2006/types DistinguishedFolderId"`
}

// --- Response ---

// --- SOAP envelope for unmarshaling ---

type findItemResponseEnvelope struct {
	XMLName xml.Name             `xml:"Envelope"`
	Header  ServerVersionInfo    `xml:"Header"`
	Body    findItemResponseBody `xml:"Body"`
}

type findItemResponseBody struct {
	FindItemResponse FindItemResponse `xml:"http://schemas.microsoft.com/exchange/services/2006/messages FindItemResponse"`
}
type FindItemResponse struct {
	ResponseMessages FindItemResponseMessages `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
}

type FindItemResponseMessages struct {
	FindItemResponseMessage FindItemResponseMessage `xml:"http://schemas.microsoft.com/exchange/services/2006/messages FindItemResponseMessage"`
}

type FindItemResponseMessage struct {
	ResponseClass ResponseClass `xml:"ResponseClass,attr"`
	ResponseCode  string        `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseCode"`
	RootFolder    RootFolder    `xml:"http://schemas.microsoft.com/exchange/services/2006/messages RootFolder"`
}

type RootFolder struct {
	TotalItemsInView        int   `xml:"TotalItemsInView,attr"`
	IncludesLastItemInRange bool  `xml:"IncludesLastItemInRange,attr"`
	Items                   Items `xml:"http://schemas.microsoft.com/exchange/services/2006/types Items"`
}

// -- Config --

type FindItemTraversal string

const (
	FindItemTraversalShallow    FindItemTraversal = "Shallow"
	FindItemTraversalAssociated FindItemTraversal = "Associated"
)

type FindItemRequestConfig struct {
	Traversal            *FindItemTraversal
	BaseShape            *BaseShape
	Query                *string
	AdditionalProperties *AdditionalProperties
	Restriction          *Restriction
}

// --- Helper function to create a FindItem request ---
func NewFindItemRequest(distinguishedFolderId DistinguishedFolderId, config FindItemRequestConfig) *FindItemRequest {
	traversal := FindItemTraversalAssociated
	if config.Traversal != nil {
		traversal = *config.Traversal
	}

	baseShape := BaseShapeAllProperties
	if config.BaseShape != nil {
		baseShape = *config.BaseShape
	}

	additionalProperties := &AdditionalProperties{}
	if config.AdditionalProperties != nil {
		additionalProperties = config.AdditionalProperties
	}

	req := &FindItemRequest{
		Traversal: string(traversal),
		ItemShape: ItemShape{
			BaseShape:            baseShape,
			AdditionalProperties: additionalProperties,
		},
		ParentFolderIds: ParentFolders{
			FolderIds: []DistinguishedFolderId{
				distinguishedFolderId,
			},
		},
	}

	if config.Restriction != nil {
		req.Restriction = config.Restriction
	}

	if config.Query != nil {
		req.QueryString = config.Query
	}

	return req
}

// --- Example function to send request ---
func FindItem(c Client, folderId string, config FindItemRequestConfig) (*FindItemResponse, error) {
	req := NewFindItemRequest(NewDistinguishedFolderId(folderId, utils.Ptr(c.GetUsername())), config)
	xmlBytes, err := xml.MarshalIndent(req, "", "  ")
	if err != nil {
		return nil, err
	}

	bb, err := c.SendAndReceive(xmlBytes)
	if err != nil {
		return nil, err
	}

	var soapResp findItemResponseEnvelope

	err = xml.Unmarshal(bb, &soapResp)
	if err != nil {
		return nil, err
	}

	return &soapResp.Body.FindItemResponse, nil
}
