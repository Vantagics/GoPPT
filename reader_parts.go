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

// readTheme reads the theme XML and extracts the color scheme and the font
// scheme. Both hang off the same part, so both are read from the same bytes:
// the colors decide what "accent1" means, the fonts decide what "+mj-lt" means,
// and a deck that uses either without the other renders in the wrong face or
// the wrong colour.
func (r *PPTXReader) readTheme(zr *zip.Reader, pres *Presentation) {
	var data []byte
	for _, path := range []string{"ppt/theme/theme1.xml", "ppt/theme/theme2.xml"} {
		b, err := readFileFromZip(zr, path)
		if err == nil {
			data = b
			break
		}
	}
	if data == nil {
		return
	}
	r.parseThemeColors(data, pres)
	r.parseThemeFonts(data, pres)
	r.parseThemeFormatScheme(data, pres)
}

// themeColorOp is a single colour transform from a theme style body — the
// children a <a:schemeClr> carries where a style says "phClr". tint and shade
// mix in linear light, lumMod/lumOff work on HLS luminance, satMod scales HLS
// saturation (a no-op on the achromatic phClr most stop lists use).
type themeColorOp struct {
	op  string
	val float64 // 0..1 (lumOff: an offset; the rest: a scale)
}

// themeGradStop is one gradient stop: a position along the gradient vector
// (0..100000) and the transforms that turn phClr into the stop colour.
type themeGradStop struct {
	pos int
	ops []themeColorOp
}

// themeFillStyle is one entry of the theme's <a:fillStyleLst>: the Office
// themes ship a solid, a subtle gradient and an intense one, and a fillRef's
// idx is a 1-based index into exactly this list.
type themeFillStyle struct {
	solid bool
	stops []themeGradStop
	angle int // gradient direction in degrees, OOXML convention
}

// themeLnStyle is one entry of <a:lnStyleLst>: a line weight and the transforms
// that turn the lnRef's colour into the line colour.
type themeLnStyle struct {
	widthEMU int
	ops      []themeColorOp
}

// themeEffectStyle is one entry of <a:effectStyleLst>: the outer shadow an
// <a:effectRef idx="N"> resolves to. The Office themes' first two styles carry
// exactly one outerShdw each (the third adds scene3d/sp3d, which no renderer
// branch walks); where a style carries no shadow the entry stays nil and the
// reference resolves to none.
type themeEffectStyle struct {
	shadow *Shadow
}

// parseThemeFormatScheme reads <a:fmtScheme>'s fill and line style lists.
// parseThemeColors stops at </a:clrScheme> — deliberately, because everything
// after it is per-index style bodies, which only a style reference consumes.
func (r *PPTXReader) parseThemeFormatScheme(data []byte, pres *Presentation) {
	decoder := xml.NewDecoder(bytes.NewReader(data))

	var (
		inFillLst, inLnLst, inGrad bool
		curFill                    *themeFillStyle
		curLn                      *themeLnStyle
		curStop                    *themeGradStop
		inSchemeClr                bool
		schemeOps                  []themeColorOp

		inEffLst, inEffShadow bool
		curEff                *themeEffectStyle
		effShadowColor        string
		effShadowAlpha        int
	)
	attachOps := func() {
		switch {
		case curStop != nil:
			curStop.ops = schemeOps
		case curLn != nil && inSchemeClr:
			curLn.ops = schemeOps
		}
		schemeOps = nil
		inSchemeClr = false
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "fillStyleLst":
				inFillLst = true
			case "lnStyleLst":
				inLnLst = true
			case "gradFill":
				if inFillLst {
					inGrad = true
					curFill = &themeFillStyle{}
				}
			case "solidFill":
				if inLnLst && curLn != nil {
					// the line colour container; ops attach to curLn
				} else if inFillLst && !inGrad {
					curFill = &themeFillStyle{solid: true}
					curFill.stops = append(curFill.stops, themeGradStop{pos: 0})
				}
			case "lin":
				if inGrad {
					for _, attr := range t.Attr {
						if attr.Name.Local == "ang" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								curFill.angle = v / 60000
							}
						}
					}
				}
			case "gs":
				if inGrad {
					curStop = &themeGradStop{}
					for _, attr := range t.Attr {
						if attr.Name.Local == "pos" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								curStop.pos = v
							}
						}
					}
				}
			case "ln":
				if inLnLst {
					curLn = &themeLnStyle{}
					for _, attr := range t.Attr {
						if attr.Name.Local == "w" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								curLn.widthEMU = v
							}
						}
					}
				}
			case "effectStyleLst":
				inEffLst = true
			case "effectStyle":
				if inEffLst {
					curEff = &themeEffectStyle{}
				}
			case "outerShdw":
				// The one effect the model can carry. The Office themes give
				// each style a single outerShdw (black, alpha'd, downward);
				// scene3d/sp3d that follow in style 3 have no counterpart and
				// are skipped by never being read.
				if curEff != nil && curEff.shadow == nil {
					inEffShadow = true
					effShadowColor = "000000"
					effShadowAlpha = 100
					sh := NewShadow()
					sh.Visible = true
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "blurRad":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								sh.BlurRadius = v / 12700
							}
						case "dist":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								sh.Distance = v / 12700
							}
						case "dir":
							if v, err := strconv.Atoi(attr.Value); err == nil {
								sh.Direction = v / 60000
							}
						}
					}
					curEff.shadow = sh
				}
			case "alpha":
				if inEffShadow {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.Atoi(attr.Value); err == nil {
								effShadowAlpha = v / 1000
							}
						}
					}
				}
			case "schemeClr", "srgbClr":
				if (inGrad && curStop != nil) || (inLnLst && curLn != nil) {
					inSchemeClr = true
					schemeOps = nil
				} else if inEffShadow {
					// The shadow's colour: usually a literal srgbClr, but a
					// schemeClr resolves through the theme the same way the
					// slide scanner resolves one.
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if argb, ok := pres.themeColors[attr.Value]; ok && argb != "" {
								effShadowColor = argb
							} else if len(attr.Value) == 6 {
								effShadowColor = attr.Value
							}
						}
					}
				}
			case "tint", "shade", "lumMod", "lumOff", "satMod":
				if inSchemeClr {
					for _, attr := range t.Attr {
						if attr.Name.Local == "val" {
							if v, err := strconv.ParseFloat(attr.Value, 64); err == nil {
								schemeOps = append(schemeOps, themeColorOp{op: t.Name.Local, val: v / 100000.0})
							}
						}
					}
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "gs":
				if inGrad && curStop != nil {
					curFill.stops = append(curFill.stops, *curStop)
					curStop = nil
				}
			case "gradFill":
				if inGrad {
					pres.themeFillStyles = append(pres.themeFillStyles, *curFill)
					inGrad = false
					curFill = nil
				}
			case "solidFill":
				if inFillLst && !inGrad && curFill != nil {
					pres.themeFillStyles = append(pres.themeFillStyles, *curFill)
					curFill = nil
				}
			case "ln":
				if inLnLst && curLn != nil {
					pres.themeLnStyles = append(pres.themeLnStyles, *curLn)
					curLn = nil
				}
			case "schemeClr", "srgbClr":
				if inSchemeClr {
					attachOps()
				}
			case "outerShdw":
				if inEffShadow && curEff != nil && curEff.shadow != nil {
					curEff.shadow.Color = NewColor(effShadowColor)
					curEff.shadow.Alpha = effShadowAlpha
					inEffShadow = false
				}
			case "effectStyle":
				if curEff != nil {
					pres.themeEffectStyles = append(pres.themeEffectStyles, *curEff)
					curEff = nil
				}
			case "effectStyleLst":
				inEffLst = false
			case "fillStyleLst":
				inFillLst = false
			case "lnStyleLst":
				inLnLst = false
			}
		}
	}
}

// clamp01 keeps an HLS→RGB channel inside 0..1 (satMod with S>1 overflows).
func clamp01(v float64) float64 {
	if v < 0 {
		return 0
	}
	if v > 1 {
		return 1
	}
	return v
}

// applyColorOps folds a transform list into a colour, in document order. The
// spaces differ per op (see applyColorTransforms): tint/shade linear,
// lumMod/lumOff/satMod on the HLS axes.
func applyColorOps(c Color, ops []themeColorOp) Color {
	for _, o := range ops {
		switch o.op {
		case "tint":
			applyTint(&c, o.val)
		case "shade":
			applyShade(&c, o.val)
		case "lumMod", "lumOff", "satMod":
			h, l, s := rgbToHLS(float64(c.GetRed())/255, float64(c.GetGreen())/255, float64(c.GetBlue())/255)
			switch o.op {
			case "lumMod":
				l *= o.val
			case "lumOff":
				l += o.val
			case "satMod":
				s *= o.val
			}
			if l > 1 {
				l = 1
			} else if l < 0 {
				l = 0
			}
			// satMod is the one op whose S PowerPoint does NOT clamp before
			// converting back: S>1 carries through and only the final RGB
			// channels saturate (ED7D31 satMod 160% → FF7200, measured). The
			// channel clamp below absorbs the overflow.
			if o.op != "satMod" {
				if s > 1 {
					s = 1
				} else if s < 0 {
					s = 0
				}
			}
			r, g, b := hlsToRGB(h, l, s)
			r = clamp01(r)
			g = clamp01(g)
			b = clamp01(b)
			c = Color{ARGB: fmtARGB(uint8(r*255+0.5), uint8(g*255+0.5), uint8(b*255+0.5))}
		}
	}
	return c
}

// resolveThemeFillStyle builds the fill a <p:style> fillRef names: idx picks a
// body from the theme's fillStyleLst, and every phClr in that body is the
// reference's scheme colour (with the slide-level transforms already applied —
// the reference colour, transforms and all, is what gets substituted). With no
// parsed theme the reference degrades to a solid of its own colour, the
// pre-r30 reading, so synthetic tests and theme-less decks keep working.
func resolveThemeFillStyle(pres *Presentation, idx int, scheme string, slideOps []themeColorOp) *Fill {
	argb, ok := "", false
	if pres != nil && pres.themeColors != nil {
		argb, ok = pres.themeColors[scheme]
	}
	if !ok || argb == "" {
		return nil
	}
	ph := applyColorOps(NewColor(argb), slideOps)
	if pres == nil || idx < 1 || idx > len(pres.themeFillStyles) {
		return NewFill().SetSolid(ph)
	}
	style := pres.themeFillStyles[idx-1]
	if len(style.stops) == 0 {
		return nil
	}
	colors := make([]Color, len(style.stops))
	for i, st := range style.stops {
		colors[i] = applyColorOps(ph, st.ops)
	}
	f := NewFill()
	if style.solid || len(colors) == 1 {
		return f.SetSolid(colors[0])
	}
	f.SetGradientLinear(colors[0], colors[len(colors)-1], style.angle)
	if len(colors) >= 3 {
		// The model carries two stops; the middle one rides along as an
		// optional knee so the renderer can reproduce the three-stop
		// subtlety gradients the Office themes define everywhere.
		f.MidColor = colors[1]
		f.MidPos = style.stops[1].pos
	}
	return f
}

// parseThemeColors populates pres.themeColors with mappings like "dk1" → "FF000000".
func (r *PPTXReader) parseThemeColors(data []byte, pres *Presentation) {
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

// parseThemeFonts populates pres.themeFonts from the theme's <a:fontScheme>:
//
//	<a:fontScheme><a:majorFont><a:latin typeface="Calibri Light"/>…
//
// A theme reference is "+mj" or "+mn" (major/minor) followed by the script it
// picks — "lt" (latin), "ea" (East Asian), "cs" (complex script). A deck that
// inherits one is the common case rather than the exception: a placeholder with
// no font of its own is drawn in the theme's major or minor face, so without
// this map every such run is drawn in whatever the font cache falls back to,
// and because the fallback measures differently, lines wrap in the wrong place.
func (r *PPTXReader) parseThemeFonts(data []byte, pres *Presentation) {
	pres.themeFonts = make(map[string]string)
	decoder := xml.NewDecoder(bytes.NewReader(data))

	var scope string // "majorFont" or "minorFont" while inside one
	inScriptFont := false

	for {
		token, err := decoder.Token()
		if err != nil {
			return
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "majorFont", "minorFont":
				scope = t.Name.Local
			case "font":
				// <a:font script="Jpan"> overrides one script inside a
				// major/minor font. Honouring it would mean tracking the
				// script of every run, so it is skipped rather than applied
				// to every script at once.
				inScriptFont = true
			case "latin", "ea", "cs":
				if scope == "" || inScriptFont {
					continue
				}
				prefix := "+mn"
				if scope == "majorFont" {
					prefix = "+mj"
				}
				for _, attr := range t.Attr {
					if attr.Name.Local != "typeface" || attr.Value == "" {
						continue
					}
					pres.themeFonts[prefix+"-"+themeFontScriptSuffix(t.Name.Local)] = attr.Value
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "font":
				inScriptFont = false
			case "majorFont", "minorFont":
				scope = ""
			case "fontScheme":
				return
			}
		}
	}
}

func themeFontScriptSuffix(local string) string {
	switch local {
	case "ea":
		return "ea"
	case "cs":
		return "cs"
	default:
		return "lt"
	}
}

// resolveThemeTypeface turns a typeface attribute into a name a font file can
// be found for, resolving <a:fontScheme> references and dropping the ones this
// deck's theme does not answer.
func resolveThemeTypeface(pres *Presentation, typeface string) string {
	if typeface == "" {
		return ""
	}
	if !strings.HasPrefix(typeface, "+") {
		return typeface
	}
	if pres == nil {
		return ""
	}
	return pres.themeFonts[typeface]
}

// typefaceOf is resolveThemeTypeface for one element's attributes: <a:latin>,
// <a:ea>, <a:cs> and <a:buFont> all carry the font as typeface="…". It returns
// "" for an element that names nothing drawable, so callers can tell "no font
// given" from "a font given" without repeating the reference test at every site.
func typefaceOf(pres *Presentation, attrs []xml.Attr) string {
	for _, attr := range attrs {
		if attr.Name.Local == "typeface" {
			return resolveThemeTypeface(pres, attr.Value)
		}
	}
	return ""
}
