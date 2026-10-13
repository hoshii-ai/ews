package ews

import (
	"context"
	"encoding/xml"
	"errors"
	"fmt"
)

// IdFormat is an EWS identifier format, as used by ConvertId.
type IdFormat string

const (
	IdFormatEwsLegacyId IdFormat = "EwsLegacyId"
	IdFormatEwsId       IdFormat = "EwsId"
	IdFormatEntryId     IdFormat = "EntryId"
	IdFormatHexEntryId  IdFormat = "HexEntryId"
	IdFormatStoreId     IdFormat = "StoreId"
	IdFormatOwaId       IdFormat = "OwaId"
)

func (f IdFormat) valid() bool {
	switch f {
	case IdFormatEwsLegacyId, IdFormatEwsId, IdFormatEntryId, IdFormatHexEntryId, IdFormatStoreId, IdFormatOwaId:
		return true
	}
	return false
}

// AlternateId is an identifier in a given format. Mailbox is the SMTP address
// of the mailbox the id belongs to; EWS requires it on every id.
type AlternateId struct {
	Format  IdFormat `xml:"Format,attr"`
	Id      string   `xml:"Id,attr"`
	Mailbox string   `xml:"Mailbox,attr"`
}

// ConvertIdResult is the outcome for one source id. Err is that id's own
// response error (e.g. ErrorInvalidIdMalformed); the other ids are still
// answered.
type ConvertIdResult struct {
	Id  AlternateId
	Err error
}

type convertIdRequest struct {
	XMLName           xml.Name `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ConvertId"`
	DestinationFormat IdFormat `xml:"DestinationFormat,attr"`
	SourceIds         struct {
		AlternateId []alternateIdXML `xml:"http://schemas.microsoft.com/exchange/services/2006/types AlternateId"`
	} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages SourceIds"`
}

type alternateIdXML struct {
	Format  IdFormat `xml:"Format,attr"`
	Id      string   `xml:"Id,attr"`
	Mailbox string   `xml:"Mailbox,attr"`
}

type convertIdMessage struct {
	responseMessage
	AlternateId *AlternateId `xml:"AlternateId"`
}

type convertIdEnvelope struct {
	Body struct {
		Response struct {
			Messages struct {
				Message []convertIdMessage `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ConvertIdResponseMessage"`
			} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ResponseMessages"`
		} `xml:"http://schemas.microsoft.com/exchange/services/2006/messages ConvertIdResponse"`
	} `xml:"Body"`
}

// ConvertId converts ids to the destination format with one request. The
// result has one entry per source id, in the same order. A request-level
// failure is returned as the error; a per-id failure is that entry's Err.
func ConvertId(ctx context.Context, c ContextClient, ids []AlternateId, destination IdFormat, opts ...RequestOption) ([]ConvertIdResult, error) {
	if len(ids) == 0 {
		return nil, errors.New("ews: ConvertId needs at least one id")
	}
	if !destination.valid() {
		return nil, fmt.Errorf("ews: ConvertId unknown destination format %q", destination)
	}
	req := convertIdRequest{DestinationFormat: destination}
	for i, id := range ids {
		if id.Id == "" || !id.Format.valid() {
			return nil, fmt.Errorf("ews: ConvertId id %d needs an id and a valid format", i)
		}
		req.SourceIds.AlternateId = append(req.SourceIds.AlternateId, alternateIdXML(id))
	}
	var env convertIdEnvelope
	if err := call(ctx, c, &req, &env, opts); err != nil {
		return nil, err
	}
	msgs := env.Body.Response.Messages.Message
	if len(msgs) != len(ids) {
		if len(msgs) == 1 {
			if err := msgs[0].err(); err != nil {
				return nil, err
			}
		}
		return nil, fmt.Errorf("ews: ConvertId asked %d ids, got %d response messages", len(ids), len(msgs))
	}
	res := make([]ConvertIdResult, len(msgs))
	for i, m := range msgs {
		if err := m.err(); err != nil {
			res[i].Err = err
			continue
		}
		if m.AlternateId == nil {
			res[i].Err = errors.New("ews: ConvertId response has no id")
			continue
		}
		res[i].Id = *m.AlternateId
	}
	return res, nil
}
