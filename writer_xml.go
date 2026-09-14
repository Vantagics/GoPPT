package gopresentation

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"time"
)

// XML namespace constants
const (
	nsRelationships  = "http://schemas.openxmlformats.org/package/2006/relationships"
	nsContentTypes   = "http://schemas.openxmlformats.org/package/2006/content-types"
	nsPresentationML = "http://schemas.openxmlformats.org/presentationml/2006/main"
	nsDrawingML      = "http://schemas.openxmlformats.org/drawingml/2006/main"
	nsOfficeDocRels  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	nsPackageRels    = "http://schemas.openxmlformats.org/package/2006/relationships"
	nsDCTerms        = "http://purl.org/dc/terms/"
	nsDC             = "http://purl.org/dc/elements/1.1/"
	nsCoreProperties = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
	nsExtProperties  = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
	nsXSI            = "http://www.w3.org/2001/XMLSchema-instance"

	// nsMarkupCompat is the namespace of mc:AlternateContent, which carries the
	// p14 version of an element alongside a plain fallback.
	nsMarkupCompat = "http://schemas.openxmlformats.org/markup-compatibility/2006"
	// nsPowerPoint2010 is the p14 namespace. A slide transition's duration
	// (p14:dur) lives here rather than in the base schema.
	nsPowerPoint2010 = "http://schemas.microsoft.com/office/powerpoint/2010/main"

	relTypeSlide       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
	relTypeSlideMaster = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
	relTypeSlideLayout = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
	relTypeTheme       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
	relTypePresProps   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"
	relTypeViewProps   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"
	relTypeTableStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
	relTypeOfficeDoc   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	relTypeCoreProps   = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
	relTypeExtProps    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
	relTypeImage       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	relTypeHyperlink   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
	relTypeChart       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
	relTypeComment     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
	relTypeCommentAuth = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
	relTypeNotesSlide  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
	relTypeNotesMaster = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
	relTypeCustomProps = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"

	// nsCustomProperties is the part-level namespace of docProps/custom.xml, and
	// nsDocPropsVT the namespace its values are typed in.
	nsCustomProperties = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
	nsDocPropsVT       = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

	// customPropsFmtid is the format id every custom property carries. It is a
	// fixed GUID chosen by the spec for "user defined property" — not something
	// this library invents — and PowerPoint rejects a part that omits it.
	customPropsFmtid = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

	// actionSlideJump is the action PowerPoint puts on an <a:hlinkClick> that
	// jumps to another slide of the same presentation. The slide itself is
	// named by the relationship the id points at, so the action and the
	// relationship only mean anything together. Written and read in one place
	// each: writer_slide.go and parseHyperlinkClick.
	actionSlideJump = "ppaction://hlinksldjump"

	ctPresentation   = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
	ctSlide          = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
	ctSlideMaster    = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
	ctSlideLayout    = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
	ctTheme          = "application/vnd.openxmlformats-officedocument.theme+xml"
	ctPresProps      = "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"
	ctViewProps      = "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"
	ctTableStyles    = "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"
	ctCoreProps      = "application/vnd.openxmlformats-package.core-properties+xml"
	ctExtProps       = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
	ctRels           = "application/vnd.openxmlformats-package.relationships+xml"
	ctChart          = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
	ctComments       = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml"
	ctCommentAuthors = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"
	ctNotesSlide     = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
	ctCustomProps    = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
)

func writeXMLToZip(zw *zip.Writer, path string, v interface{}) error {
	fw, err := zw.Create(path)
	if err != nil {
		return fmt.Errorf("failed to create %s in zip: %w", path, err)
	}
	if _, err := fw.Write([]byte(xml.Header)); err != nil {
		return err
	}
	enc := xml.NewEncoder(fw)
	enc.Indent("", "  ")
	if err := enc.Encode(v); err != nil {
		return fmt.Errorf("failed to encode %s: %w", path, err)
	}
	return nil
}

func writeRawXMLToZip(zw *zip.Writer, path string, content string) error {
	fw, err := zw.Create(path)
	if err != nil {
		return fmt.Errorf("failed to create %s in zip: %w", path, err)
	}
	_, err = fw.Write([]byte(content))
	return err
}

// --- Content Types ---

type xmlContentTypes struct {
	XMLName   xml.Name      `xml:"Types"`
	Xmlns     string        `xml:"xmlns,attr"`
	Defaults  []xmlDefault  `xml:"Default"`
	Overrides []xmlOverride `xml:"Override"`
}

type xmlDefault struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

type xmlOverride struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

func (w *PPTXWriter) writeContentTypes(zw *zip.Writer) error {
	ct := xmlContentTypes{
		Xmlns: nsContentTypes,
		Defaults: []xmlDefault{
			{Extension: "rels", ContentType: ctRels},
			{Extension: "xml", ContentType: "application/xml"},
		},
		Overrides: []xmlOverride{
			{PartName: "/ppt/presentation.xml", ContentType: ctPresentation},
			{PartName: "/ppt/presProps.xml", ContentType: ctPresProps},
			{PartName: "/ppt/viewProps.xml", ContentType: ctViewProps},
			{PartName: "/ppt/tableStyles.xml", ContentType: ctTableStyles},
			{PartName: "/ppt/slideMasters/slideMaster1.xml", ContentType: ctSlideMaster},
			{PartName: "/ppt/slideLayouts/slideLayout1.xml", ContentType: ctSlideLayout},
			{PartName: "/ppt/theme/theme1.xml", ContentType: ctTheme},
			{PartName: "/docProps/core.xml", ContentType: ctCoreProps},
			{PartName: "/docProps/app.xml", ContentType: ctExtProps},
		},
	}

	// docProps/custom.xml only exists when the presentation defines custom
	// properties: a declared part the package does not contain is what makes
	// PowerPoint ask to repair the file.
	if w.hasCustomProperties() {
		ct.Overrides = append(ct.Overrides, xmlOverride{
			PartName:    "/docProps/custom.xml",
			ContentType: ctCustomProps,
		})
	}

	// Add slide content types
	for i := range w.presentation.slides {
		ct.Overrides = append(ct.Overrides, xmlOverride{
			PartName:    fmt.Sprintf("/ppt/slides/slide%d.xml", i+1),
			ContentType: ctSlide,
		})
	}

	// Add image defaults
	for _, slide := range w.presentation.slides {
		for _, ds := range collectDrawingShapes(slide.shapes) {
			ext := w.getImageExtension(ds)
			found := false
			for _, d := range ct.Defaults {
				if d.Extension == ext {
					found = true
					break
				}
			}
			if !found {
				ct.Defaults = append(ct.Defaults, xmlDefault{
					Extension:   ext,
					ContentType: w.getImageContentType(ds),
				})
			}
		}
	}

	// Add chart content types
	chartIdx := 1
	for _, slide := range w.presentation.slides {
		for _, shape := range slide.shapes {
			if _, ok := shape.(*ChartShape); ok {
				ct.Overrides = append(ct.Overrides, xmlOverride{
					PartName:    fmt.Sprintf("/ppt/charts/chart%d.xml", chartIdx),
					ContentType: ctChart,
				})
				chartIdx++
			}
		}
	}

	// Add comment content types
	if w.hasComments() {
		ct.Overrides = append(ct.Overrides, xmlOverride{
			PartName:    "/ppt/commentAuthors.xml",
			ContentType: ctCommentAuthors,
		})
		for i, slide := range w.presentation.slides {
			if len(slide.comments) > 0 {
				ct.Overrides = append(ct.Overrides, xmlOverride{
					PartName:    fmt.Sprintf("/ppt/comments/comment%d.xml", i+1),
					ContentType: ctComments,
				})
			}
		}
	}

	// Add notes slide content types
	for i, slide := range w.presentation.slides {
		if slide.notes != "" {
			ct.Overrides = append(ct.Overrides, xmlOverride{
				PartName:    fmt.Sprintf("/ppt/notesSlides/notesSlide%d.xml", i+1),
				ContentType: ctNotesSlide,
			})
		}
	}

	return writeXMLToZip(zw, "[Content_Types].xml", ct)
}

func (w *PPTXWriter) getImageExtension(ds *DrawingShape) string {
	if ds.mimeType != "" {
		switch ds.mimeType {
		case "image/png":
			return "png"
		case "image/jpeg":
			return "jpeg"
		case "image/gif":
			return "gif"
		case "image/bmp":
			return "bmp"
		case "image/svg+xml":
			return "svg"
		}
	}
	if ds.path != "" {
		ext := strings.TrimPrefix(filepath.Ext(ds.path), ".")
		if ext == "jpg" {
			return "jpeg"
		}
		return ext
	}
	return "png"
}

func (w *PPTXWriter) getImageContentType(ds *DrawingShape) string {
	if ds.mimeType != "" {
		return ds.mimeType
	}
	ext := w.getImageExtension(ds)
	switch ext {
	case "png":
		return "image/png"
	case "jpeg", "jpg":
		return "image/jpeg"
	case "gif":
		return "image/gif"
	case "bmp":
		return "image/bmp"
	case "svg":
		return "image/svg+xml"
	default:
		return "image/png"
	}
}

// --- Relationships ---

type xmlRelationships struct {
	XMLName       xml.Name          `xml:"Relationships"`
	Xmlns         string            `xml:"xmlns,attr"`
	Relationships []xmlRelationship `xml:"Relationship"`
}

type xmlRelationship struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

func (w *PPTXWriter) writeRootRels(zw *zip.Writer) error {
	rels := xmlRelationships{
		Xmlns: nsRelationships,
		Relationships: []xmlRelationship{
			{ID: "rId1", Type: relTypeOfficeDoc, Target: "ppt/presentation.xml"},
			{ID: "rId2", Type: relTypeCoreProps, Target: "docProps/core.xml"},
			{ID: "rId3", Type: relTypeExtProps, Target: "docProps/app.xml"},
		},
	}
	if w.hasCustomProperties() {
		rels.Relationships = append(rels.Relationships, xmlRelationship{
			ID:     "rId4",
			Type:   relTypeCustomProps,
			Target: "docProps/custom.xml",
		})
	}
	return writeXMLToZip(zw, "_rels/.rels", rels)
}

func (w *PPTXWriter) writePresentationRels(zw *zip.Writer) error {
	rels := xmlRelationships{
		Xmlns: nsRelationships,
	}

	relIdx := 1
	// Slide master
	rels.Relationships = append(rels.Relationships, xmlRelationship{
		ID:     fmt.Sprintf("rId%d", relIdx),
		Type:   relTypeSlideMaster,
		Target: "slideMasters/slideMaster1.xml",
	})
	relIdx++

	// Slides
	for i := range w.presentation.slides {
		rels.Relationships = append(rels.Relationships, xmlRelationship{
			ID:     fmt.Sprintf("rId%d", relIdx),
			Type:   relTypeSlide,
			Target: fmt.Sprintf("slides/slide%d.xml", i+1),
		})
		relIdx++
	}

	// PresProps
	rels.Relationships = append(rels.Relationships, xmlRelationship{
		ID:     fmt.Sprintf("rId%d", relIdx),
		Type:   relTypePresProps,
		Target: "presProps.xml",
	})
	relIdx++

	// ViewProps
	rels.Relationships = append(rels.Relationships, xmlRelationship{
		ID:     fmt.Sprintf("rId%d", relIdx),
		Type:   relTypeViewProps,
		Target: "viewProps.xml",
	})
	relIdx++

	// TableStyles
	rels.Relationships = append(rels.Relationships, xmlRelationship{
		ID:     fmt.Sprintf("rId%d", relIdx),
		Type:   relTypeTableStyles,
		Target: "tableStyles.xml",
	})
	relIdx++

	// Theme
	rels.Relationships = append(rels.Relationships, xmlRelationship{
		ID:     fmt.Sprintf("rId%d", relIdx),
		Type:   relTypeTheme,
		Target: "theme/theme1.xml",
	})
	relIdx++

	// Comment authors
	if w.hasComments() {
		rels.Relationships = append(rels.Relationships, xmlRelationship{
			ID:     fmt.Sprintf("rId%d", relIdx),
			Type:   relTypeCommentAuth,
			Target: "commentAuthors.xml",
		})
	}

	return writeXMLToZip(zw, "ppt/_rels/presentation.xml.rels", rels)
}

// --- App Properties ---

func (w *PPTXWriter) writeAppProperties(zw *zip.Writer) error {
	props := w.presentation.properties
	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="%s" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>GoPresentation v%s</Application>
  <Company>%s</Company>
  <AppVersion>%s</AppVersion>
  <Slides>%d</Slides>
</Properties>`, nsExtProperties, Version, xmlEscape(props.Company), Version, len(w.presentation.slides))
	return writeRawXMLToZip(zw, "docProps/app.xml", content)
}

// --- Core Properties ---

func (w *PPTXWriter) writeCoreProperties(zw *zip.Writer) error {
	props := w.presentation.properties
	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="%s" xmlns:dc="%s" xmlns:dcterms="%s" xmlns:xsi="%s">
  <dc:creator>%s</dc:creator>
  <cp:lastModifiedBy>%s</cp:lastModifiedBy>
  <dc:title>%s</dc:title>
  <dc:description>%s</dc:description>
  <dc:subject>%s</dc:subject>
  <cp:keywords>%s</cp:keywords>
  <cp:category>%s</cp:category>
  <cp:revision>%s</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>
</cp:coreProperties>`,
		nsCoreProperties, nsDC, nsDCTerms, nsXSI,
		xmlEscape(props.Creator),
		xmlEscape(props.LastModifiedBy),
		xmlEscape(props.Title),
		xmlEscape(props.Description),
		xmlEscape(props.Subject),
		xmlEscape(props.Keywords),
		xmlEscape(props.Category),
		xmlEscape(props.Revision),
		props.Created.UTC().Format("2006-01-02T15:04:05Z"),
		props.Modified.UTC().Format("2006-01-02T15:04:05Z"),
	)
	return writeRawXMLToZip(zw, "docProps/core.xml", content)
}

// --- Custom Properties ---

// hasCustomProperties reports whether the presentation defines any custom
// document property. It decides whether docProps/custom.xml is written and
// declared at all: a declared part the package does not contain, or a part the
// package contains but does not declare, is what makes PowerPoint offer to
// repair the file.
func (w *PPTXWriter) hasCustomProperties() bool {
	props := w.presentation.properties
	return props != nil && len(props.customProps) > 0
}

// writeCustomProperties writes docProps/custom.xml.
//
// API.md documents SetCustomProperty beside GetCustomPropertyValue as one pair,
// and neither half reached the file: no code wrote the part at all. A custom
// property was readable back through the in-memory API and gone the moment the
// presentation was saved.
//
// The names are sorted because the model holds them in a map. Walking a map
// emits the properties in a different order on every save, which makes a
// package differ from itself for reasons that have nothing to do with what the
// caller changed.
func (w *PPTXWriter) writeCustomProperties(zw *zip.Writer) error {
	if !w.hasCustomProperties() {
		return nil
	}
	props := w.presentation.properties

	names := make([]string, 0, len(props.customProps))
	for name := range props.customProps {
		names = append(names, name)
	}
	sort.Strings(names)

	var b strings.Builder
	for i, name := range names {
		// pid 1 is reserved by the spec; the first user property is 2.
		b.WriteString(fmt.Sprintf("\n  <property fmtid=\"%s\" pid=\"%d\" name=\"%s\">%s</property>",
			customPropsFmtid, i+2, xmlEscape(name), customPropertyValueXML(props.customProps[name])))
	}

	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="%s" xmlns:vt="%s">%s
</Properties>`, nsCustomProperties, nsDocPropsVT, b.String())
	return writeRawXMLToZip(zw, "docProps/custom.xml", content)
}

// customPropertyValueXML types a custom property value the way its declared
// PropertyType says.
//
// A value whose Go type does not match its declared one is written through its
// string form rather than dropped: a custom property is the caller's data, and
// losing it silently is worse than writing it in the least specific type.
func customPropertyValueXML(prop *CustomProperty) string {
	if prop == nil {
		return "<vt:lpwstr></vt:lpwstr>"
	}
	text := xmlEscape(fmt.Sprintf("%v", prop.Value))

	switch prop.Type {
	case PropertyTypeBoolean:
		v := false
		switch value := prop.Value.(type) {
		case bool:
			v = value
		case string:
			v, _ = strconv.ParseBool(value)
		}
		return fmt.Sprintf("<vt:bool>%t</vt:bool>", v)
	case PropertyTypeInteger:
		if v, err := strconv.ParseInt(fmt.Sprintf("%v", prop.Value), 10, 64); err == nil {
			return fmt.Sprintf("<vt:i4>%d</vt:i4>", v)
		}
	case PropertyTypeFloat:
		if v, err := strconv.ParseFloat(fmt.Sprintf("%v", prop.Value), 64); err == nil {
			return fmt.Sprintf("<vt:r8>%g</vt:r8>", v)
		}
	case PropertyTypeDate:
		if v, ok := prop.Value.(time.Time); ok {
			return fmt.Sprintf("<vt:filetime>%s</vt:filetime>", v.UTC().Format("2006-01-02T15:04:05Z"))
		}
	}
	return fmt.Sprintf("<vt:lpwstr>%s</vt:lpwstr>", text)
}

// xmlEscape escapes special XML characters using the standard library.
func xmlEscape(s string) string {
	var b strings.Builder
	if err := xml.EscapeText(&b, []byte(s)); err != nil {
		// EscapeText writing to strings.Builder never fails, but handle gracefully.
		return s
	}
	return b.String()
}

// colorRGB safely extracts the 6-character RGB portion from an 8-character ARGB string.
// Returns "000000" if the input is invalid.
func colorRGB(c Color) string {
	if len(c.ARGB) >= 8 {
		return c.ARGB[2:]
	}
	if len(c.ARGB) == 6 {
		return c.ARGB
	}
	return "000000"
}
