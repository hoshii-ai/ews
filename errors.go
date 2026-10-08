package ews

import (
	"fmt"
	"strconv"
	"strings"
)

// Common EWS ResponseCode values.
const (
	ResponseCodeNoError                   = "NoError"
	ResponseCodeErrorServerBusy           = "ErrorServerBusy"
	ResponseCodeErrorSubscriptionNotFound = "ErrorSubscriptionNotFound"
)

// ResponseError is an EWS response message with ResponseClass=Error (or a
// SOAP fault carrying an EWS ResponseCode).
type ResponseError struct {
	Class       ResponseClass
	Code        string // ResponseCode, e.g. ErrorServerBusy
	MessageText string
	// BackOffMilliseconds is set for ErrorServerBusy (0 if absent).
	BackOffMilliseconds int
	// SubscriptionIds is the ErrorSubscriptionIds list of GetStreamingEvents errors.
	SubscriptionIds []string
}

func (e *ResponseError) Error() string {
	if e.MessageText == "" {
		return e.Code
	}
	return fmt.Sprintf("%s: %s", e.Code, e.MessageText)
}

// IsServerBusy reports whether the error is ErrorServerBusy.
func (e *ResponseError) IsServerBusy() bool { return e.Code == ResponseCodeErrorServerBusy }

func backOff(values []MessageXmlValue) int {
	for _, v := range values {
		if v.Name == "BackOffMilliseconds" {
			n, err := strconv.Atoi(strings.TrimSpace(v.Value))
			if err == nil {
				return n
			}
		}
	}
	return 0
}

// subscriptionIdList is a list of <t:SubscriptionId> elements.
type subscriptionIdList struct {
	Ids []string `xml:"http://schemas.microsoft.com/exchange/services/2006/types SubscriptionId"`
}

// responseMessage holds the fields common to every EWS response message.
// Embed it in operation-specific message structs.
type responseMessage struct {
	ResponseClass        ResponseClass      `xml:"ResponseClass,attr"`
	ResponseCode         string             `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseCode"`
	MessageText          string             `xml:"http://schemas.microsoft.com/exchange/services/2006/messages MessageText"`
	MessageXml           MessageXml         `xml:"http://schemas.microsoft.com/exchange/services/2006/messages MessageXml"`
	ErrorSubscriptionIds subscriptionIdList `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ErrorSubscriptionIds"`
}

// err returns a *ResponseError if the message is an error, else nil.
func (m responseMessage) err() error {
	if m.ResponseClass != ResponseClassError {
		return nil
	}
	return &ResponseError{
		Class:               m.ResponseClass,
		Code:                m.ResponseCode,
		MessageText:         m.MessageText,
		BackOffMilliseconds: backOff(m.MessageXml.Values),
		SubscriptionIds:     m.ErrorSubscriptionIds.Ids,
	}
}

type serverVersionHeader struct {
	ServerVersionInfo ServerVersionInfo `xml:"ServerVersionInfo"`
}

// FolderRef identifies a folder either by distinguished name or by id.
// Set DistinguishedId (inbox, sentitems, ...) or Id (+ ChangeKey).
type FolderRef struct {
	DistinguishedId string
	Id              string
	ChangeKey       string
}

func (f FolderRef) isZero() bool { return f.DistinguishedId == "" && f.Id == "" }

type folderRefs struct {
	DistinguishedFolderId []distinguishedRef `xml:"http://schemas.microsoft.com/exchange/services/2006/types DistinguishedFolderId,omitempty"`
	FolderId              []folderIdRef      `xml:"http://schemas.microsoft.com/exchange/services/2006/types FolderId,omitempty"`
}

type distinguishedRef struct {
	Id      string      `xml:"Id,attr"`
	Mailbox *mailboxRef `xml:"http://schemas.microsoft.com/exchange/services/2006/types Mailbox,omitempty"`
}

// mailboxRef addresses another mailbox (delegate access).
type mailboxRef struct {
	EmailAddress string `xml:"http://schemas.microsoft.com/exchange/services/2006/types EmailAddress"`
}

type folderIdRef struct {
	Id        string `xml:"Id,attr"`
	ChangeKey string `xml:"ChangeKey,attr,omitempty"`
}

func newFolderRefs(refs []FolderRef) folderRefs {
	var r folderRefs
	for _, f := range refs {
		if f.DistinguishedId != "" {
			r.DistinguishedFolderId = append(r.DistinguishedFolderId, distinguishedRef{Id: f.DistinguishedId})
		} else {
			r.FolderId = append(r.FolderId, folderIdRef{Id: f.Id, ChangeKey: f.ChangeKey})
		}
	}
	return r
}
