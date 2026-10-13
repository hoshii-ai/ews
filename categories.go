package ews

import (
	"bytes"
	"encoding/base64"
	"encoding/xml"
	"errors"
	"fmt"
	"strings"
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
	ItemId           ItemId     `xml:"-"` // This itemID is not present in the XML. We just store it here to make it easier to update the item.
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
		if strings.EqualFold(c.Name, name) {
			return fmt.Errorf("category with name %q already exists", name)
		}
	}

	// Generate GUID
	newGuid := fmt.Sprintf("{%s}", uuid.New().String())

	// Now time
	now := time.Now().UTC()

	// Append
	never := categoryTimeNever
	usageCount := uint(0)
	renameOnFirstUse := 0
	cl.Categories = append(cl.Categories, Category{
		Name:                 name,
		Color:                color,
		KeyboardShortcut:     0, // None
		UsageCount:           &usageCount,
		LastTimeUsedNotes:    &never,
		LastTimeUsedJournal:  &never,
		LastTimeUsedContacts: &never,
		LastTimeUsedTasks:    &never,
		LastTimeUsedCalendar: &never,
		LastTimeUsedMail:     &never,
		LastTimeUsed:         now,
		LastSessionUsed:      cl.LastSavedSession + 1,
		GUID:                 newGuid,
		RenameOnFirstUse:     &renameOnFirstUse,
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
		if strings.EqualFold(c.Name, name) {
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

// Outlook writes and expects [MS-OXOCFG] dateTimes: UTC, millisecond precision, no zone designator.
const categoryTimeLayout = "2006-01-02T15:04:05.000"

var categoryTimeNever = time.Date(1601, 1, 1, 0, 0, 0, 0, time.UTC)

func formatCategoryTime(t time.Time) string {
	return t.UTC().Format(categoryTimeLayout)
}

func formatOptionalCategoryTime(t *time.Time) string {
	if t == nil {
		return ""
	}
	return formatCategoryTime(*t)
}

// parseCategoryTime accepts what Outlook writes (no zone, UTC, any fraction of a second)
// and RFC 3339. An empty value is the zero time.
func parseCategoryTime(s string) (time.Time, error) {
	if s == "" {
		return time.Time{}, nil
	}
	if t, err := time.ParseInLocation("2006-01-02T15:04:05.999999999", s, time.UTC); err == nil {
		return t, nil
	}
	return time.Parse(time.RFC3339Nano, s)
}

func parseOptionalCategoryTime(s string) (*time.Time, error) {
	if s == "" {
		return nil, nil
	}
	t, err := parseCategoryTime(s)
	if err != nil {
		return nil, err
	}
	return &t, nil
}

type categoryListWire struct {
	XMLName          xml.Name       `xml:"categories"`
	Default          string         `xml:"default,attr"`
	LastSavedSession int            `xml:"lastSavedSession,attr"`
	LastSavedTime    string         `xml:"lastSavedTime,attr"`
	Xmlns            string         `xml:"xmlns,attr"`
	Categories       []categoryWire `xml:"category"`
}

type categoryWire struct {
	Name                 string    `xml:"name,attr"`
	Color                ColorType `xml:"color,attr"`
	KeyboardShortcut     uint      `xml:"keyboardShortcut,attr"`
	UsageCount           *uint     `xml:"usageCount,attr,omitempty"`
	LastTimeUsedNotes    string    `xml:"lastTimeUsedNotes,attr,omitempty"`
	LastTimeUsedJournal  string    `xml:"lastTimeUsedJournal,attr,omitempty"`
	LastTimeUsedContacts string    `xml:"lastTimeUsedContacts,attr,omitempty"`
	LastTimeUsedTasks    string    `xml:"lastTimeUsedTasks,attr,omitempty"`
	LastTimeUsedCalendar string    `xml:"lastTimeUsedCalendar,attr,omitempty"`
	LastTimeUsedMail     string    `xml:"lastTimeUsedMail,attr,omitempty"`
	LastTimeUsed         string    `xml:"lastTimeUsed,attr"`
	LastSessionUsed      int       `xml:"lastSessionUsed,attr"`
	GUID                 string    `xml:"guid,attr"`
	RenameOnFirstUse     *int      `xml:"renameOnFirstUse,attr,omitempty"`
}

func (cl CategoryList) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	wire := categoryListWire{
		Default:          cl.Default,
		LastSavedSession: cl.LastSavedSession,
		LastSavedTime:    formatCategoryTime(cl.LastSavedTime),
		Xmlns:            xlmns,
		Categories:       make([]categoryWire, 0, len(cl.Categories)),
	}
	for _, c := range cl.Categories {
		wire.Categories = append(wire.Categories, categoryWire{
			Name:                 c.Name,
			Color:                c.Color,
			KeyboardShortcut:     c.KeyboardShortcut,
			UsageCount:           c.UsageCount,
			LastTimeUsedNotes:    formatOptionalCategoryTime(c.LastTimeUsedNotes),
			LastTimeUsedJournal:  formatOptionalCategoryTime(c.LastTimeUsedJournal),
			LastTimeUsedContacts: formatOptionalCategoryTime(c.LastTimeUsedContacts),
			LastTimeUsedTasks:    formatOptionalCategoryTime(c.LastTimeUsedTasks),
			LastTimeUsedCalendar: formatOptionalCategoryTime(c.LastTimeUsedCalendar),
			LastTimeUsedMail:     formatOptionalCategoryTime(c.LastTimeUsedMail),
			LastTimeUsed:         formatCategoryTime(c.LastTimeUsed),
			LastSessionUsed:      c.LastSessionUsed,
			GUID:                 c.GUID,
			RenameOnFirstUse:     c.RenameOnFirstUse,
		})
	}
	return e.Encode(wire)
}

func (cl *CategoryList) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var wire categoryListWire
	if err := d.DecodeElement(&wire, &start); err != nil {
		return err
	}
	lastSavedTime, err := parseCategoryTime(wire.LastSavedTime)
	if err != nil {
		return fmt.Errorf("lastSavedTime: %w", err)
	}
	*cl = CategoryList{
		ItemId:           cl.ItemId,
		XMLName:          wire.XMLName,
		Xmlns:            xlmns,
		Default:          wire.Default,
		LastSavedSession: wire.LastSavedSession,
		LastSavedTime:    lastSavedTime,
		Categories:       make([]Category, 0, len(wire.Categories)),
	}
	seenGUIDs := make(map[string]struct{}, len(wire.Categories))
	seenNames := make(map[string]struct{}, len(wire.Categories))
	for _, w := range wire.Categories {
		// Outlook stops rendering colours when the list repeats an entry, so keep the first of each.
		guid, name := strings.ToLower(w.GUID), strings.ToLower(w.Name)
		if _, dup := seenGUIDs[guid]; dup && guid != "" {
			continue
		}
		if _, dup := seenNames[name]; dup {
			continue
		}
		seenGUIDs[guid] = struct{}{}
		seenNames[name] = struct{}{}
		c := Category{
			Name:             w.Name,
			Color:            w.Color,
			KeyboardShortcut: w.KeyboardShortcut,
			UsageCount:       w.UsageCount,
			LastSessionUsed:  w.LastSessionUsed,
			GUID:             w.GUID,
			RenameOnFirstUse: w.RenameOnFirstUse,
		}
		if c.LastTimeUsed, err = parseCategoryTime(w.LastTimeUsed); err != nil {
			return fmt.Errorf("category %q lastTimeUsed: %w", w.Name, err)
		}
		for _, f := range []struct {
			dst **time.Time
			src string
		}{
			{&c.LastTimeUsedNotes, w.LastTimeUsedNotes},
			{&c.LastTimeUsedJournal, w.LastTimeUsedJournal},
			{&c.LastTimeUsedContacts, w.LastTimeUsedContacts},
			{&c.LastTimeUsedTasks, w.LastTimeUsedTasks},
			{&c.LastTimeUsedCalendar, w.LastTimeUsedCalendar},
			{&c.LastTimeUsedMail, w.LastTimeUsedMail},
		} {
			if *f.dst, err = parseOptionalCategoryTime(f.src); err != nil {
				return fmt.Errorf("category %q: %w", w.Name, err)
			}
		}
		cl.Categories = append(cl.Categories, c)
	}
	return nil
}
