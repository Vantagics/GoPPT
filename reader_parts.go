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
