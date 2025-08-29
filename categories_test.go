package ews

import (
	"fmt"
	"testing"
	"time"
)

func Test_Categories(t *testing.T) {
	// Initialize a category list
	cl := CategoryList{
		Default:          "Blue",
		LastSavedSession: 1,
		LastSavedTime:    time.Now().UTC(),
	}
	xmlStr, _ := cl.ToXML()
	fmt.Println("Initial XML:")
	fmt.Println(xmlStr)

	// Add new category
	if err := cl.AddCategory("Green", ColorGreen); err != nil {
		panic(err)
	}
	if err := cl.AddCategory("Blue", ColorBlue); err != nil {
		panic(err)
	}

	xmlStr, _ = cl.ToXML()
	fmt.Println("After adding Green and Blue:")
	fmt.Println(xmlStr)

	// Delete an old category
	if err := cl.DeleteCategory("Blue"); err != nil {
		fmt.Println("Delete error:", err)
	}

	// Marshal back to XML
	xmlStr, _ = cl.ToXML()
	fmt.Println("Final XML:")
	fmt.Println(xmlStr)
}
