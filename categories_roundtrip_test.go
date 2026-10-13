package ews

import (
	"encoding/base64"
	"strings"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// Real-shaped blob: Outlook/OWA write UTC times without a zone designator.
const outlookCategoryList = `<?xml version="1.0" encoding="utf-8"?>
<categories default="Blue" lastSavedSession="3" lastSavedTime="2026-10-13T13:45:49.180" xmlns="CategoryList.xsd">
  <category name="Blue" color="7" keyboardShortcut="0" usageCount="2" lastTimeUsedMail="2026-10-12T08:01:02.5" lastTimeUsed="2026-10-12T08:01:02.500" lastSessionUsed="2" guid="{11111111-1111-1111-1111-111111111111}"/>
  <category name="Green" color="4" keyboardShortcut="0" lastTimeUsed="2026-10-11T00:00:00" lastSessionUsed="1" guid="{22222222-2222-2222-2222-222222222222}"/>
</categories>`

func TestCategoryListZonelessTimesRoundTrip(t *testing.T) {
	cl, err := CategoryListFromBase64(base64.StdEncoding.EncodeToString([]byte(outlookCategoryList)))
	require.NoError(t, err)
	assert.Equal(t, time.Date(2026, 10, 13, 13, 45, 49, 180_000_000, time.UTC), cl.LastSavedTime)
	require.Len(t, cl.Categories, 2)
	blue, green := cl.Categories[0], cl.Categories[1]
	assert.Equal(t, time.Date(2026, 10, 12, 8, 1, 2, 500_000_000, time.UTC), blue.LastTimeUsed)
	require.NotNil(t, blue.LastTimeUsedMail)
	assert.Equal(t, blue.LastTimeUsed, *blue.LastTimeUsedMail)
	assert.Nil(t, blue.LastTimeUsedNotes)
	assert.Equal(t, time.Date(2026, 10, 11, 0, 0, 0, 0, time.UTC), green.LastTimeUsed)
	assert.Nil(t, green.LastTimeUsedMail)

	require.NoError(t, cl.AddCategory("Red", ColorRed))
	b64, err := cl.CategoryListToBase64()
	require.NoError(t, err)
	raw, err := base64.StdEncoding.DecodeString(b64)
	require.NoError(t, err)
	out := string(raw)

	// Outlook's layout: zone-less, millisecond precision. Optional attributes stay absent when unset.
	assert.Contains(t, out, `lastSavedTime="`)
	assert.NotContains(t, out, "Z\"")
	assert.Contains(t, out, `lastTimeUsedMail="2026-10-12T08:01:02.500"`)
	assert.Contains(t, out, `lastTimeUsed="2026-10-11T00:00:00.000"`)
	assert.Equal(t, 1, strings.Count(out, "lastTimeUsedMail=\"2026-10-12"))

	// Re-decoding what we wrote keeps the untouched values.
	again, err := CategoryListFromBase64(b64)
	require.NoError(t, err)
	require.Len(t, again.Categories, 3)
	assert.Equal(t, blue.LastTimeUsed, again.Categories[0].LastTimeUsed)
	assert.Equal(t, blue.LastTimeUsedMail, again.Categories[0].LastTimeUsedMail)
}

func TestCategoryListTimeParsing(t *testing.T) {
	for in, want := range map[string]time.Time{
		"":                            {},
		"2026-10-13T13:45:49":         time.Date(2026, 10, 13, 13, 45, 49, 0, time.UTC),
		"2026-10-13T13:45:49.180":     time.Date(2026, 10, 13, 13, 45, 49, 180_000_000, time.UTC),
		"2026-10-13T13:45:49.1234567": time.Date(2026, 10, 13, 13, 45, 49, 123_456_700, time.UTC),
		"2026-10-13T13:45:49Z":        time.Date(2026, 10, 13, 13, 45, 49, 0, time.UTC),
	} {
		got, err := parseCategoryTime(in)
		require.NoError(t, err, in)
		assert.True(t, want.Equal(got), in)
	}
	_, err := parseCategoryTime("yesterday")
	assert.Error(t, err)
}
