package ews

import (
	"fmt"
	"time"
)

type ResponseClass string

const (
	ResponseClassSuccess ResponseClass = "Success"
	ResponseClassWarning ResponseClass = "Warning"
	ResponseClassError   ResponseClass = "Error"
)

type Response struct {
	ResponseClass ResponseClass `xml:"ResponseClass,attr"`
	MessageText   string        `xml:"http://schemas.microsoft.com/exchange/services/2006/messages MessageText"`
	ResponseCode  string        `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseCode"`
	MessageXml    MessageXml    `xml:"http://schemas.microsoft.com/exchange/services/2006/messages MessageXml"`
	Items         Items         `xml:"http://schemas.microsoft.com/exchange/services/2006/messages Items"`
}

type ServerVersionInfo struct {
	MajorVersion     string `xml:"MajorVersion,attr"`
	MinorVersion     string `xml:"MinorVersion,attr"`
	MajorBuildNumber string `xml:"MajorBuildNumber,attr"`
	MinorBuildNumber string `xml:"MinorBuildNumber,attr"`
	Version          string `xml:"Version,attr"`
	XmlnsH           string `xml:"xmlns:h,attr,omitempty"`
	Xmlns            string `xml:"xmlns,attr,omitempty"`
	XmlnsXsd         string `xml:"xmlns:xsd,attr,omitempty"`
	XmlnsXsi         string `xml:"xmlns:xsi,attr,omitempty"`
}

type EmailAddress struct {
	Name         string `xml:"http://schemas.microsoft.com/exchange/services/2006/types Name"`
	EmailAddress string `xml:"http://schemas.microsoft.com/exchange/services/2006/types EmailAddress"`
	RoutingType  string `xml:"http://schemas.microsoft.com/exchange/services/2006/types RoutingType"`
	MailboxType  string `xml:"http://schemas.microsoft.com/exchange/services/2006/types MailboxType"`
}

type ItemId struct {
	Id        string `xml:"Id,attr"`
	ChangeKey string `xml:"ChangeKey,attr,omitempty"`
}

type AttachmentId struct {
	Id string `xml:"Id,attr"`
}

type MessageXml struct {
	ExceptionType       string `xml:"ExceptionType"`
	ExceptionCode       string `xml:"ExceptionCode"`
	ExceptionServerName string `xml:"ExceptionServerName"`
	ExceptionMessage    string `xml:"ExceptionMessage"`
}

type DistinguishedFolderId struct {
	// List of values:
	// https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/distinguishedfolderid
	Id                 string   `xml:"Id,attr"`
	Mailbox            *Mailbox `xml:"http://schemas.microsoft.com/exchange/services/2006/types Mailbox,omitempty"`
	IncludeHiddenItems *bool    `xml:"http://schemas.microsoft.com/exchange/services/2006/types IncludeHiddenItems,omitempty"`
}

func NewDistinguishedFolderId(id string, email *string) DistinguishedFolderId {
	var mailbox *Mailbox
	if email != nil {
		mailbox = &Mailbox{
			EmailAddress: *email,
		}
	}

	return DistinguishedFolderId{
		Id:      id,
		Mailbox: mailbox,
	}
}

type ParentFolderId struct {
	DistinguishedFolderId DistinguishedFolderId `xml:"http://schemas.microsoft.com/exchange/services/2006/types DistinguishedFolderId"`
}

type Persona struct {
	PersonaId            PersonaId            `xml:"http://schemas.microsoft.com/exchange/services/2006/types PersonaId"`
	DisplayName          string               `xml:"http://schemas.microsoft.com/exchange/services/2006/types DisplayName"`
	Title                string               `xml:"http://schemas.microsoft.com/exchange/services/2006/types Title"`
	Department           string               `xml:"http://schemas.microsoft.com/exchange/services/2006/types Department"`
	Departments          Departments          `xml:"http://schemas.microsoft.com/exchange/services/2006/types Departments"`
	EmailAddress         EmailAddress         `xml:"http://schemas.microsoft.com/exchange/services/2006/types EmailAddress"`
	RelevanceScore       int                  `xml:"http://schemas.microsoft.com/exchange/services/2006/types RelevanceScore"`
	BusinessPhoneNumbers BusinessPhoneNumbers `xml:"http://schemas.microsoft.com/exchange/services/2006/types BusinessPhoneNumbers"`
	MobilePhones         MobilePhones         `xml:"http://schemas.microsoft.com/exchange/services/2006/types MobilePhones"`
	OfficeLocations      OfficeLocations      `xml:"http://schemas.microsoft.com/exchange/services/2006/types OfficeLocations"`
}

type PersonaId struct {
	Id string `xml:"Id,attr"`
}

type BusinessPhoneNumbers struct {
	PhoneNumberAttributedValue PhoneNumberAttributedValue `xml:"http://schemas.microsoft.com/exchange/services/2006/types PhoneNumberAttributedValue"`
}

type MobilePhones struct {
	PhoneNumberAttributedValue PhoneNumberAttributedValue `xml:"http://schemas.microsoft.com/exchange/services/2006/types PhoneNumberAttributedValue"`
}

type Value struct {
	Number string `json:"Number"`
	Type   string `json:"Type"`
}

type PhoneNumberAttributedValue struct {
	Value Value `json:"Value"`
}

type OfficeLocations struct {
	StringAttributedValue StringAttributedValue `xml:"StringAttributedValue"`
}

type Departments struct {
	StringAttributedValue StringAttributedValue `xml:"StringAttributedValue"`
}

type StringAttributedValue struct {
	Value string `json:"Value"`
}

type FieldURI struct {
	// List of possible values:
	// https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/fielduri
	FieldURI string `xml:"FieldURI,attr,omitempty"`
}

type (
	PropertyType string
	PropertyTag  string
)

const (
	PropertyTypeBinary   PropertyType = "Binary"
	PropertyTypeString   PropertyType = "String"
	PropertyTypeInteger  PropertyType = "Integer"
	PropertyTypeBoolean  PropertyType = "Boolean"
	PropertyTypeDateTime PropertyType = "DateTime"
	PropertyTypeDouble   PropertyType = "Double"
	PropertyTypeSingle   PropertyType = "Single"
)

const (
	PropertyTagCategories        PropertyTag = "0x7c08"
	PropertyTagInternetMessageId PropertyTag = "0x1035"
)

type ExtendedFieldURI struct {
	PropertyTag  PropertyTag  `xml:"PropertyTag,attr,omitempty"`
	PropertyType PropertyType `xml:"PropertyType,attr,omitempty"`
	PropertyName string       `xml:"PropertyName,attr,omitempty"`
	PropertyId   string       `xml:"PropertyId,attr,omitempty"`
}

type BaseShape string

const (
	BaseShapeDefault       BaseShape = "Default"
	BaseShapeIdOnly        BaseShape = "IdOnly"
	BaseShapeAllProperties BaseShape = "AllProperties"
)

type ItemShape struct {
	BaseShape                 BaseShape             `xml:"http://schemas.microsoft.com/exchange/services/2006/types BaseShape"`
	IncludeMimeContent        bool                  `xml:"http://schemas.microsoft.com/exchange/services/2006/types IncludeMimeContent,omitempty"`
	BodyType                  string                `xml:"http://schemas.microsoft.com/exchange/services/2006/types BodyType,omitempty"`
	FilterHtmlContent         bool                  `xml:"http://schemas.microsoft.com/exchange/services/2006/types FilterHtmlContent,omitempty"`
	ConvertHtmlCodePageToUTF8 bool                  `xml:"http://schemas.microsoft.com/exchange/services/2006/types ConvertHtmlCodePageToUTF8,omitempty"`
	AdditionalProperties      *AdditionalProperties `xml:"http://schemas.microsoft.com/exchange/services/2006/types AdditionalProperties,omitempty"`
}

type AdditionalProperties struct {
	FieldURI         []FieldURI         `xml:"http://schemas.microsoft.com/exchange/services/2006/types FieldURI,omitempty"`
	ExtendedFieldURI []ExtendedFieldURI `xml:"http://schemas.microsoft.com/exchange/services/2006/types ExtendedFieldURI,omitempty"`
	// add additional fields
}
type Time string

func (t Time) ToTime() (time.Time, error) {
	offset, err := getRFC3339Offset(time.Now())
	if err != nil {
		return time.Time{}, err
	}
	return time.Parse(time.RFC3339, string(t)+offset)

}

// return RFC3339 formatted offset, ex: +03:00 -03:30
func getRFC3339Offset(t time.Time) (string, error) {

	_, offset := t.Zone()
	i := int(float32(offset) / 36)

	sign := "+"
	if i < 0 {
		i = -i
		sign = "-"
	}
	hour := i / 100
	min := i % 100
	min = (60 * min) / 100

	return fmt.Sprintf("%s%02d:%02d", sign, hour, min), nil
}
