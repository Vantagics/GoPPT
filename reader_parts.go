package gopresentation

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
	"time"
)

// --- Core Properties ---

func (r *PPTXReader) readCoreProperties(zr *zip.Reader, pres *Presentation) error {
	data, err := readFileFromZip(zr, "docProps/core.xml")
	if err != nil {
		return err
	}

	decoder := xml.NewDecoder(bytes.NewReader(data))
	props := pres.properties

	var currentElement string
	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			currentElement = t.Name.Local
		case xml.CharData:
			text := strings.TrimSpace(string(t))
			if text == "" {
				continue
			}
			switch currentElement {
			case "creator":
				props.Creator = text
			case "lastModifiedBy":
				props.LastModifiedBy = text
			case "title":
				props.Title = text
			case "description":
				props.Description = text
			case "subject":
				props.Subject = text
			case "keywords":
				props.Keywords = text
			case "category":
				props.Category = text
			case "revision":
				props.Revision = text
			case "created":
				if t, err := time.Parse("2006-01-02T15:04:05Z", text); err == nil {
					props.Created = t
				}
			case "modified":
				if t, err := time.Parse("2006-01-02T15:04:05Z", text); err == nil {
					props.Modified = t
				}
			}
		}
	}
	return nil
}

// readAppProperties reads docProps/app.xml.
//
// The writer has always emitted this part and nothing read it back, so a
// company name reached the file and then quietly reverted to the default on the
// next open. Only the fields the writer writes are read: Application and
// AppVersion name the producer, not the document.
func (r *PPTXReader) readAppProperties(zr *zip.Reader, pres *Presentation) error {
	data, err := readFileFromZip(zr, "docProps/app.xml")
	if err != nil {
		return err
	}
	props := pres.properties
	if props == nil {
		return nil
	}

	decoder := xml.NewDecoder(bytes.NewReader(data))
	var currentElement string
	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := token.(type) {
		case xml.StartElement:
			currentElement = t.Name.Local
		case xml.CharData:
			if currentElement != "Company" {
				continue
			}
			if text := strings.TrimSpace(string(t)); text != "" {
				props.Company = text
			}
		}
	}
	return nil
}

// --- Custom Properties ---

// readCustomProperties reads docProps/custom.xml back into the document
// properties.
//
// A custom property is typed by the name of a <vt:*> child element rather than
// by an attribute, and keyed by a name attribute on <property>. A value whose
// type this library has no PropertyType for is kept as a string with its text
// intact, so nothing in the part is silently dropped.
func (r *PPTXReader) readCustomProperties(zr *zip.Reader, pres *Presentation) error {
	data, err := readFileFromZip(zr, "docProps/custom.xml")
	if err != nil {
		return err
	}
	props := pres.properties
	if props == nil {
		return nil
	}
	if props.customProps == nil {
		props.customProps = make(map[string]*CustomProperty)
	}

	decoder := xml.NewDecoder(bytes.NewReader(data))
	// openKind gates the character data — it is the <vt:*> element currently
	// being read, and is cleared when that element closes. lastKind remembers
	// which one it was, because the value is committed when <property> closes,
	// by which point the value element has already been closed and openKind is
	// back to being empty.
	var (
		name     string
		openKind string
		lastKind string
		text     string
		inProp   bool
	)
	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "property":
				inProp = true
				name, openKind, lastKind, text = "", "", "", ""
				for _, attr := range t.Attr {
					if attr.Name.Local == "name" {
						name = attr.Value
					}
				}
			default:
				if inProp && isCustomPropertyValueElement(t.Name.Local) {
					openKind, lastKind = t.Name.Local, t.Name.Local
				}
			}
		case xml.CharData:
			if inProp && openKind != "" {
				text += string(t)
			}
		case xml.EndElement:
			if t.Name.Local == "property" {
				if inProp && name != "" {
					props.customProps[name] = parseCustomProperty(name, lastKind, strings.TrimSpace(text))
				}
				inProp = false
			} else if inProp && t.Name.Local == openKind {
				openKind = ""
			}
		}
	}
	return nil
}

// isCustomPropertyValueElement reports whether a local element name is one of
// the docPropsVTypes value elements.
func isCustomPropertyValueElement(local string) bool {
	switch local {
	case "lpwstr", "lpstr", "bstr", "i1", "i2", "i4", "i8", "int", "uint",
		"ui1", "ui2", "ui4", "ui8", "r4", "r8", "decimal", "bool", "filetime", "date":
		return true
	}
	return false
}

// parseCustomProperty builds a CustomProperty from the text of one <vt:*>
// element. A value that does not parse as its declared type keeps its text and
// is reported as a string.
func parseCustomProperty(name, kind, text string) *CustomProperty {
	prop := &CustomProperty{Name: name, Value: text, Type: PropertyTypeString}
	switch kind {
	case "bool":
		if v, err := strconv.ParseBool(text); err == nil {
			prop.Type = PropertyTypeBoolean
			prop.Value = v
		}
	case "i1", "i2", "i4", "i8", "int", "uint", "ui1", "ui2", "ui4", "ui8":
		if v, err := strconv.ParseInt(text, 10, 64); err == nil {
			prop.Type = PropertyTypeInteger
			prop.Value = v
		}
	case "r4", "r8", "decimal":
		if v, err := strconv.ParseFloat(text, 64); err == nil {
			prop.Type = PropertyTypeFloat
			prop.Value = v
		}
	case "filetime", "date":
		if v, err := time.Parse("2006-01-02T15:04:05Z", text); err == nil {
			prop.Type = PropertyTypeDate
			prop.Value = v
		}
	}
	return prop
}

// --- Presentation ---

type xmlPresentation struct {
	XMLName        xml.Name          `xml:"presentation"`
	SldMasterIdLst xmlSldMasterIdLst `xml:"sldMasterIdLst"`
	SldIdLst       xmlSldIdLst       `xml:"sldIdLst"`
	SldSz          xmlSldSz          `xml:"sldSz"`
	NotesSz        xmlNotesSz        `xml:"notesSz"`
}

type xmlSldMasterIdLst struct {
	SldMasterIds []xmlSldMasterId `xml:"sldMasterId"`
}

type xmlSldMasterId struct {
	ID string `xml:"id,attr"`
}

type xmlSldIdLst struct {
	SldIds []xmlSldId `xml:"sldId"`
}

type xmlSldId struct {
	ID string `xml:"id,attr"`
}

type xmlSldSz struct {
	CX   string `xml:"cx,attr"`
	CY   string `xml:"cy,attr"`
	Type string `xml:"type,attr"`
}

type xmlNotesSz struct {
	CX string `xml:"cx,attr"`
	CY string `xml:"cy,attr"`
}

func (r *PPTXReader) readPresentation(zr *zip.Reader, pres *Presentation) ([]string, error) {
	data, err := readFileFromZip(zr, "ppt/presentation.xml")
	if err != nil {
		return nil, fmt.Errorf("failed to read presentation.xml: %w", err)
	}

	// Parse using streaming to handle namespaces properly
	decoder := xml.NewDecoder(bytes.NewReader(data))
	var slideRelIDs []string

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "sldSz":
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "cx":
						if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							pres.layout.CX = v
						}
					case "cy":
						if v, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							pres.layout.CY = v
						}
					case "type":
						pres.layout.Name = attr.Value
					}
				}
			case "sldId":
				for _, attr := range t.Attr {
					if attr.Name.Local == "id" && attr.Name.Space != "" {
						slideRelIDs = append(slideRelIDs, attr.Value)
					} else if attr.Name.Local == "id" && attr.Name.Space == "" {
						// This is the numeric ID, not the relationship ID
					}
				}
			}
		}
	}

	// If we didn't find relationship IDs via namespace, try reading rels directly
	if len(slideRelIDs) == 0 {
		rels, err := r.readRelationships(zr, "ppt/_rels/presentation.xml.rels")
		if err == nil {
			for _, rel := range rels {
				if rel.Type == relTypeSlide {
					slideRelIDs = append(slideRelIDs, rel.ID)
				}
			}
		}
	}

	return slideRelIDs, nil
}

// --- Theme Colors ---

// readThemeColors reads the theme XML and extracts the color scheme.
// It populates pres.themeColors with mappings like "dk1" → "FF000000".
func (r *PPTXReader) readThemeColors(zr *zip.Reader, pres *Presentation) {
	// Try common theme paths
	var data []byte
	var err error
	for _, path := range []string{"ppt/theme/theme1.xml", "ppt/theme/theme2.xml"} {
		data, err = readFileFromZip(zr, path)
		if err == nil {
			break
		}
	}
	if data == nil {
		return
	}

	pres.themeColors = make(map[string]string)
	decoder := xml.NewDecoder(bytes.NewReader(data))

	// Track which scheme color element we're inside
	var currentSchemeColor string

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "dk1", "dk2", "lt1", "lt2",
				"accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
				"hlink", "folHlink":
				currentSchemeColor = t.Name.Local
			case "srgbClr":
				if currentSchemeColor != "" {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							pres.themeColors[currentSchemeColor] = "FF" + strings.ToUpper(attr.Value)
						}
					}
				}
			case "sysClr":
				if currentSchemeColor != "" {
					for _, attr := range t.Attr {
						if attr.Name.Local == "lastClr" {
							pres.themeColors[currentSchemeColor] = "FF" + strings.ToUpper(attr.Value)
						}
					}
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "dk1", "dk2", "lt1", "lt2",
				"accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
				"hlink", "folHlink":
				currentSchemeColor = ""
			case "clrScheme":
				// Also add common aliases
				if c, ok := pres.themeColors["dk1"]; ok {
					pres.themeColors["tx1"] = c
				}
				if c, ok := pres.themeColors["lt1"]; ok {
					pres.themeColors["bg1"] = c
				}
				if c, ok := pres.themeColors["dk2"]; ok {
					pres.themeColors["tx2"] = c
				}
				if c, ok := pres.themeColors["lt2"]; ok {
					pres.themeColors["bg2"] = c
				}
				return // done
			}
		}
	}
}
