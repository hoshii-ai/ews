package ews

import (
	"encoding/xml"
	"errors"
)

type BasePoint string

const (
	BasePointBeginning BasePoint = "Beginning"
	BasePointEnd       BasePoint = "End"
)

type FindPeopleRequest struct {
	XMLName             struct{}            `xml:"m:FindPeople"`
	PersonaShape        *PersonaShape       `xml:"m:PersonaShape,omitempty"`
	IndexedPageItemView IndexedPageItemView `xml:"m:IndexedPageItemView"`
	ParentFolderId      ParentFolderId      `xml:"m:ParentFolderId"`
	QueryString         string              `xml:"m:QueryString,omitempty"`
	// add additional fields
}

type PersonaShape struct {
	BaseShape            BaseShape            `xml:"t:BaseShape,omitempty"`
	AdditionalProperties AdditionalProperties `xml:"t:AdditionalProperties,omitempty"`
}

type IndexedPageItemView struct {
	MaxEntriesReturned int       `xml:"MaxEntriesReturned,attr,omitempty"`
	Offset             int       `xml:"Offset,attr"`
	BasePoint          BasePoint `xml:"BasePoint,attr"`
}

type findPeopleResponseEnvelop struct {
	XMLName struct{}               `xml:"Envelope"`
	Body    findPeopleResponseBody `xml:"Body"`
}
type findPeopleResponseBody struct {
	FindPeopleResponse FindPeopleResponse `xml:"FindPeopleResponse"`
}

type FindPeopleResponse struct {
	Response
	People                    People `xml:"People"`
	TotalNumberOfPeopleInView int    `xml:"TotalNumberOfPeopleInView"`
	FirstMatchingRowIndex     int    `xml:"FirstMatchingRowIndex"`
	FirstLoadedRowIndex       int    `xml:"FirstLoadedRowIndex"`
}

type People struct {
	Persona []Persona `xml:"Persona"`
}

// GetUserAvailability
// https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/findpeople-operation
func FindPeople(c Client, r *FindPeopleRequest) (*FindPeopleResponse, error) {

	xmlBytes, err := xml.MarshalIndent(r, "", "  ")
	if err != nil {
		return nil, err
	}

	bb, err := c.SendAndReceive(xmlBytes)
	if err != nil {
		return nil, err
	}

	var soapResp findPeopleResponseEnvelop
	err = xml.Unmarshal(bb, &soapResp)
	if err != nil {
		return nil, err
	}

	if soapResp.Body.FindPeopleResponse.ResponseClass == ResponseClassError {
		return nil, errors.New(soapResp.Body.FindPeopleResponse.MessageText)
	}

	return &soapResp.Body.FindPeopleResponse, nil
}
