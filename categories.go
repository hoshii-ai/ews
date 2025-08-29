package ews

import (
	"bytes"
	"encoding/base64"
	"encoding/xml"
	"errors"
	"fmt"
	"time"

	"github.com/google/uuid"
)

var (
	xlmns = "CategoryList.xsd"
)

type ColorType int

const (
	ColorNone       ColorType = -1
	ColorRed        ColorType = 0
	ColorOrange     ColorType = 1
	ColorPeach      ColorType = 2
	ColorYellow     ColorType = 3
	ColorGreen      ColorType = 4
	ColorTeal       ColorType = 5
	ColorOlive      ColorType = 6
	ColorBlue       ColorType = 7
	ColorPurple     ColorType = 8
	ColorMaroon     ColorType = 9
	ColorSteel      ColorType = 10
	ColorDarkSteel  ColorType = 11
	ColorGray       ColorType = 12
	ColorDarkGray   ColorType = 13
	ColorBlack      ColorType = 14
	ColorDarkRed    ColorType = 15
	ColorDarkOrange ColorType = 16
	ColorDarkPeach  ColorType = 17
	ColorDarkYellow ColorType = 18
	ColorDarkGreen  ColorType = 19
	ColorDarkTeal   ColorType = 20
	ColorDarkOlive  ColorType = 21
	ColorDarkBlue   ColorType = 22
	ColorDarkPurple ColorType = 23
	ColorDarkMaroon ColorType = 24
)

// CategoryList represents the <categories> root element
type CategoryList struct {
	ItemId           ItemId     // This itemID is not present in the XML. We just store it here to make it easier to update the item.
	XMLName          xml.Name   `xml:"categories"`
	Xmlns            string     `xml:"xmlns,attr"`
	Categories       []Category `xml:"category"`
	Default          string     `xml:"default,attr"`
	LastSavedSession int        `xml:"lastSavedSession,attr"`
	LastSavedTime    time.Time  `xml:"lastSavedTime,attr"`
}

// Category represents a single <category> entry
type Category struct {
	Name                 string     `xml:"name,attr"`
	Color                ColorType  `xml:"color,attr"`
	KeyboardShortcut     uint       `xml:"keyboardShortcut,attr"`
	UsageCount           *uint      `xml:"usageCount,attr,omitempty"`
	LastTimeUsedNotes    *time.Time `xml:"lastTimeUsedNotes,attr,omitempty"`
	LastTimeUsedJournal  *time.Time `xml:"lastTimeUsedJournal,attr,omitempty"`
	LastTimeUsedContacts *time.Time `xml:"lastTimeUsedContacts,attr,omitempty"`
	LastTimeUsedTasks    *time.Time `xml:"lastTimeUsedTasks,attr,omitempty"`
	LastTimeUsedCalendar *time.Time `xml:"lastTimeUsedCalendar,attr,omitempty"`
	LastTimeUsedMail     *time.Time `xml:"lastTimeUsedMail,attr,omitempty"`
	LastTimeUsed         time.Time  `xml:"lastTimeUsed,attr"`
	LastSessionUsed      int        `xml:"lastSessionUsed,attr"`
	GUID                 string     `xml:"guid,attr"`
	RenameOnFirstUse     *int       `xml:"renameOnFirstUse,attr,omitempty"`
}

// AddCategory adds a new category, ensuring uniqueness of name.
func (cl *CategoryList) AddCategory(name string, color ColorType) error {
	// Check uniqueness
	for _, c := range cl.Categories {
		if c.Name == name {
			return fmt.Errorf("category with name %q already exists", name)
		}
	}

	// Generate GUID
	newGuid := fmt.Sprintf("{%s}", uuid.New().String())

	// Now time
	now := time.Now().UTC()

	// Append
	cl.Categories = append(cl.Categories, Category{
		Name:             name,
		Color:            color,
		KeyboardShortcut: 0, // None
		LastTimeUsed:     now,
		LastSessionUsed:  cl.LastSavedSession + 1,
		GUID:             newGuid,
	})

	// Update list metadata
	cl.LastSavedSession++
	cl.LastSavedTime = now

	return nil
}

// DeleteCategory removes a category by name
func (cl *CategoryList) DeleteCategory(name string) error {
	idx := -1
	for i, c := range cl.Categories {
		if c.Name == name {
			idx = i
			break
		}
	}
	if idx == -1 {
		return errors.New("category not found")
	}

	cl.Categories = append(cl.Categories[:idx], cl.Categories[idx+1:]...)

	// Update metadata
	cl.LastSavedSession++
	cl.LastSavedTime = time.Now().UTC()

	return nil
}

// ToXML marshals the CategoryList back into XML string
func (cl *CategoryList) ToXML() (string, error) {
	cl.Xmlns = xlmns
	output, err := xml.MarshalIndent(cl, "", "  ")
	if err != nil {
		return "", err
	}
	return xml.Header + string(output), nil
}

// ToBase64 marshals the CategoryList to XML and encodes it as base64
func (cl *CategoryList) CategoryListToBase64() (string, error) {
	buf := &bytes.Buffer{}
	buf.WriteString(xml.Header) // Adds <?xml version="1.0" encoding="UTF-8"?>

	enc := xml.NewEncoder(buf)
	enc.Indent("", "  ")
	if err := enc.Encode(cl); err != nil {
		return "", err
	}
	if err := enc.Flush(); err != nil {
		return "", err
	}

	return base64.StdEncoding.EncodeToString(buf.Bytes()), nil
}

// FromBase64 decodes a base64 string into a CategoryList struct
func CategoryListFromBase64(b64 string) (*CategoryList, error) {
	data, err := base64.StdEncoding.DecodeString(b64)
	if err != nil {
		return nil, err
	}

	var cl CategoryList
	if err := xml.Unmarshal(data, &cl); err != nil {
		return nil, err
	}

	return &cl, nil
}
