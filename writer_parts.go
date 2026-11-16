package gopresentation

import (
	"archive/zip"
	"fmt"
	"strings"
)

// --- Presentation Part ---

func (w *PPTXWriter) writePresentation(zw *zip.Writer) error {
	layout := w.presentation.layout

	// One sldMasterId per master part: a deck carries one master per theme
	// and every slide resolves through its layout to its own master.
	masterList := ""
	nMasters := len(w.masterPlan)
	if nMasters == 0 {
		nMasters = 1
	}
	// Master ids live above the layout id band: sldMasterId and sldLayoutId
	// ids must be unique across the whole package, and the layout lists
	// claim 2147483648+index below 1000.
	for m := 1; m <= nMasters; m++ {
		masterList += fmt.Sprintf("    <p:sldMasterId id=\"%d\" r:id=\"rId%d\"/>\n", 2147483648+1000+m, m)
	}

	slideList := ""
	relIdx := nMasters + 1 // rId1..nMasters are the slideMasters
	for i := range w.presentation.slides {
		slideList += fmt.Sprintf(`    <p:sldId id="%d" r:id="rId%d"/>
`, 256+i, relIdx)
		relIdx++
	}

	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">
  <p:sldMasterIdLst>
%s  </p:sldMasterIdLst>
  <p:sldIdLst>
%s  </p:sldIdLst>
  <p:sldSz cx="%d" cy="%d" type="%s"/>
  <p:notesSz cx="%d" cy="%d"/>
  <p:defaultTextStyle/>
</p:presentation>`,
		nsDrawingML, nsOfficeDocRels, nsPresentationML,
		masterList, slideList,
		layout.CX, layout.CY, layout.Name,
		layout.CY, layout.CX, // notes are rotated
	)
	return writeRawXMLToZip(zw, "ppt/presentation.xml", content)
}

// --- Presentation Properties ---

func (w *PPTXWriter) writePresProps(zw *zip.Writer) error {
	pp := w.presentation.presentationProperties

	showType := ""
	switch pp.slideshowType {
	case SlideshowTypePresent:
		showType = `<p:present/>`
	case SlideshowTypeBrowse:
		showType = `<p:browse/>`
	case SlideshowTypeKiosk:
		showType = `<p:kiosk/>`
	}

	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">
  <p:showPr>
    %s
  </p:showPr>
</p:presentationPr>`, nsDrawingML, nsOfficeDocRels, nsPresentationML, showType)
	return writeRawXMLToZip(zw, "ppt/presProps.xml", content)
}

// --- View Properties ---

func (w *PPTXWriter) writeViewProps(zw *zip.Writer) error {
	pp := w.presentation.presentationProperties
	lastView := "sldView"
	switch pp.lastView {
	case ViewNotes:
		lastView = "notesView"
	case ViewHandout:
		lastView = "handoutView"
	case ViewOutline:
		lastView = "outlineView"
	case ViewSlideMaster:
		lastView = "sldMasterView"
	case ViewSlideSorter:
		lastView = "sldSorterView"
	}

	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr xmlns:a="%s" xmlns:r="%s" xmlns:p="%s" lastView="%s">
  <p:normalViewPr>
    <p:restoredLeft sz="15620"/>
    <p:restoredTop sz="94660"/>
  </p:normalViewPr>
  <p:slideViewPr>
    <p:cSldViewPr>
      <p:cViewPr>
        <p:scale>
          <a:sx n="%d" d="100"/>
          <a:sy n="%d" d="100"/>
        </p:scale>
        <p:origin x="0" y="0"/>
      </p:cViewPr>
    </p:cSldViewPr>
  </p:slideViewPr>
</p:viewPr>`, nsDrawingML, nsOfficeDocRels, nsPresentationML, lastView,
		int(pp.zoom*100), int(pp.zoom*100))
	return writeRawXMLToZip(zw, "ppt/viewProps.xml", content)
}

// --- Table Styles ---

func (w *PPTXWriter) writeTableStyles(zw *zip.Writer) error {
	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="%s" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`, nsDrawingML)
	return writeRawXMLToZip(zw, "ppt/tableStyles.xml", content)
}

// --- Slide Master ---

func (w *PPTXWriter) writeSlideMaster(zw *zip.Writer) error {
	masters := w.masterPlan
	if len(masters) == 0 {
		masters = []masterPlan{{index: 1}}
	}
	for _, mp := range masters {
		if err := w.writeSlideMasterN(zw, mp); err != nil {
			return err
		}
	}
	return nil
}

// writeSlideMasterN emits one master part for the layouts grouped under it:
// the master's own clrMap (a master can send bg1 at dk2 — the whole deck's
// background turns dark), its decoration shapes, one sldLayoutId per layout,
// and its rels wiring the layouts plus the master's own theme.
func (w *PPTXWriter) writeSlideMasterN(zw *zip.Writer, mp masterPlan) error {
	st := w.presentation.masterStates[mp.masterPath]

	// The colour mapping: without the source's mapping Office resolves the
	// written master's bg1 against the built-in lt1/dk1 defaults and the
	// mapped decks' backgrounds come out wrong.
	clrMap := `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>`
	if cm, ok := w.presentation.masterClrMapCache[mp.masterPath]; ok && cm != "" {
		clrMap = cm
	}

	// The master's own background: an explicit <p:bgPr> fill rides back as
	// one; otherwise the template's bgRef 1001 (solid bg1) stands in.
	bgXML := "    <p:bg>\n      <p:bgRef idx=\"1001\">\n        <a:schemeClr val=\"bg1\"/>\n      </p:bgRef>\n    </p:bg>\n"
	if st != nil && st.background != nil && st.background.Type != FillNone && st.background.Type != FillPicture {
		bgXML = fmt.Sprintf("    <p:bg>\n      <p:bgPr>\n%s        <a:effectLst/>\n      </p:bgPr>\n    </p:bg>\n", w.writeFillXML(st.background))
	}

	// The master's decoration shapes — the furniture every slide of this
	// master draws — plus their media. relBase skips the layout rels.
	shapeID := 2
	var masterDeco []Shape
	if st != nil {
		masterDeco = st.decoShapes
	}
	decoXML, mediaRels, err := w.emitDecoShapes(zw, masterDeco, fmt.Sprintf("imageM%d", mp.index), len(mp.entries)+1, &shapeID)
	if err != nil {
		return err
	}

	layoutList := ""
	masterRels := ""
	for i, e := range mp.entries {
		layoutList += fmt.Sprintf("    <p:sldLayoutId id=\"%d\" r:id=\"rId%d\"/>\n", 2147483648+e.index, i+1)
		masterRels += fmt.Sprintf("  <Relationship Id=\"rId%d\" Type=\"%s\" Target=\"../slideLayouts/slideLayout%d.xml\"/>\n", i+1, relTypeSlideLayout, e.index)
	}

	// The theme rel lands after the layouts and the media.
	themeRel := len(mp.entries) + strings.Count(mediaRels, "<Relationship ") + 1

	// The master's text-style ladder rides back verbatim: slide placeholders
	// inherit cap, bullet and font defaults from it (the Trek master's
	// titleStyle cap="all" capitalises its slides' titles).
	txStylesXML := ""
	if st != nil {
		if tx := w.presentation.masterTxStylesCache[mp.masterPath]; tx != "" {
			txStylesXML = "  " + tx + "\n"
		}
	}

	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">
  <p:cSld>
%s    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
%s    </p:spTree>
  </p:cSld>
  %s
  <p:sldLayoutIdLst>
%s  </p:sldLayoutIdLst>
%s</p:sldMaster>`, nsDrawingML, nsOfficeDocRels, nsPresentationML, bgXML, decoXML, clrMap, layoutList, txStylesXML)

	if err := writeRawXMLToZip(zw, fmt.Sprintf("ppt/slideMasters/slideMaster%d.xml", mp.index), content); err != nil {
		return err
	}

	// Slide master rels: every layout of this master, then the media, then
	// the theme.
	rels := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="%s">
%s%s  <Relationship Id="rId%d" Type="%s" Target="../theme/theme%d.xml"/>
</Relationships>`, nsRelationships, masterRels, mediaRels, themeRel, relTypeTheme, mp.index)
	return writeRawXMLToZip(zw, fmt.Sprintf("ppt/slideMasters/_rels/slideMaster%d.xml.rels", mp.index), rels)
}

// --- Slide Layout ---

func (w *PPTXWriter) writeSlideLayout(zw *zip.Writer) error {
	return w.writeSlideLayoutN(zw, w.layoutPlan[0])
}

// isTextPlaceholder reports whether the placeholder type is one of the
// text prompts Office never paints from a layout: titles, bodies and the
// footer trio. A non-text placeholder (picture, table, chart) with explicit
// paint renders as itself.
func isTextPlaceholder(t PlaceholderType) bool {
	switch t {
	case PlaceholderTitle, PlaceholderBody, PlaceholderCtrTitle, PlaceholderSubTitle,
		PlaceholderDate, PlaceholderFooter, PlaceholderSlideNum:
		return true
	}
	return false
}

// emitDecoShapes serialises the decoration shapes of a master or layout
// part: the shapes' XML body plus the pictures' media written beside the
// part, returned as the relationship entries the part's own rels must
// carry. relBase is the first relationship number the media may use — a
// layout's rId1 goes to its master, a master's rIds go to its layouts
// first.
func (w *PPTXWriter) emitDecoShapes(zw *zip.Writer, shapes []Shape, mediaPrefix string, relBase int, shapeID *int) (string, string, error) {
	if len(shapes) == 0 {
		return "", "", nil
	}
	var body strings.Builder
	rels := strings.Builder{}
	mediaIdx := 0
	for _, sh := range shapes {
		switch s := sh.(type) {
		case *DrawingShape:
			if s.data != nil {
				// A deco picture: its media is written beside the part and
				// referenced from the part's own rels.
				mediaIdx++
				mediaName := fmt.Sprintf("%s_%d.%s", mediaPrefix, mediaIdx, w.getImageExtension(s))
				fw, ferr := zw.Create("ppt/media/" + mediaName)
				if ferr != nil {
					return "", "", ferr
				}
				if _, werr := fw.Write(s.data); werr != nil {
					return "", "", werr
				}
				relID := fmt.Sprintf("rId%d", relBase+mediaIdx-1)
				rels.WriteString(fmt.Sprintf("  <Relationship Id=\"%s\" Type=\"%s\" Target=\"../media/%s\"/>\n", relID, relTypeImage, mediaName))
				body.WriteString(w.layoutPicXML(s, shapeID, relID))
				continue
			}
			body.WriteString(w.writeDrawingShapeXML(s, shapeID, 1))
		case *AutoShape:
			body.WriteString(w.writeAutoShapeXML(s, shapeID))
		case *RichTextShape:
			body.WriteString(w.writeRichTextShapeXML(s, shapeID))
		case *LineShape:
			body.WriteString(w.writeLineShapeXML(s, shapeID))
		case *GroupShape:
			body.WriteString(w.writeGroupShapeXML(s, shapeID, 1))
		case *PlaceholderShape:
			// A painted master/layout TEXT placeholder is written back as a
			// REAL placeholder (the <p:ph> element): Office treats those as
			// edit-time prompts — it paints them only where a slide
			// instantiates one, and the slide inherits the paint natively.
			// Written as a plain shape it painted on every slide (the themes
			// deck's L23 went 1.3%→92% that way); written as nothing the
			// inherited paint was lost (the Austin master's white title card,
			// slide4 75%).
			//
			// A NON-text placeholder with paint (a picture placeholder drawn
			// with a custGeom fill — the custgeom-nofill-path deck's yellow
			// marker) renders as itself in the original: Office paints it
			// whether or not a slide instantiates it, so it goes back as a
			// plain shape.
			if isTextPlaceholder(s.phType) {
				body.WriteString(w.writePlaceholderShapeXML(s, shapeID))
			} else {
				body.WriteString(w.writeRichTextShapeXML(&s.RichTextShape, shapeID))
			}
		}
	}
	return body.String(), rels.String(), nil
}

// writeSlideLayoutN emits one layout part with its decoration shapes — the
// layout deco is where decks like KEY02 keep their whole visual identity —
// and writes the deco pictures' media alongside, wiring the layout's own
// rels to them.
func (w *PPTXWriter) writeSlideLayoutN(zw *zip.Writer, entry layoutPlanEntry) error {
	// The layout's own background: read decks like the themes corpus keep a
	// per-layout <p:bg><p:bgPr> whose fill differs per layout — without it
	// every layout round-trips onto the master's white.
	bgXML := ""
	if entry.bgFill != nil && entry.bgFill.Type != FillPicture && entry.bgFill.Type != FillNone {
		bgXML = fmt.Sprintf("    <p:bg>\n      <p:bgPr>\n%s      <a:effectLst/>\n      </p:bgPr>\n    </p:bg>\n", w.writeFillXML(entry.bgFill))
	}

	shapeID := 2
	decoXML, mediaRels, err := w.emitDecoShapes(zw, entry.decoShapes, fmt.Sprintf("imageL%d", entry.index), 2, &shapeID)
	if err != nil {
		return err
	}

	// The layout's own suppression of the master graphics rides on the root.
	suppress := ""
	if entry.hideMasterSp {
		suppress = ` showMasterSp="0"`
	}

	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"%s type="blank" preserve="1">
  <p:cSld name="Blank">
%s    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
%s    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sldLayout>`, nsDrawingML, nsOfficeDocRels, nsPresentationML, suppress, bgXML, decoXML)

	if err := writeRawXMLToZip(zw, fmt.Sprintf("ppt/slideLayouts/slideLayout%d.xml", entry.index), content); err != nil {
		return err
	}

	// Slide layout rels: the layout's own master, then the deco pictures'
	// media.
	masterIdx := entry.masterIdx
	if masterIdx == 0 {
		masterIdx = 1
	}
	rels := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="%s">
  <Relationship Id="rId1" Type="%s" Target="../slideMasters/slideMaster%d.xml"/>
%s</Relationships>`, nsRelationships, relTypeSlideMaster, masterIdx, mediaRels)
	return writeRawXMLToZip(zw, fmt.Sprintf("ppt/slideLayouts/_rels/slideLayout%d.xml.rels", entry.index), rels)
}

// layoutPicXML serialises one decoration picture of a layout part: the same
// picture body writeDrawingShapeXML emits, with the relationship id the
// layout's own rels assign.
func (w *PPTXWriter) layoutPicXML(s *DrawingShape, shapeID *int, relID string) string {
	id := *shapeID
	*shapeID++

	name := s.name
	if name == "" {
		name = fmt.Sprintf("Picture %d", id)
	}
	shadowXML := shapeEffectXML(s.shadow, s.softEdgeRad, "          ")
	prst := s.presetGeom
	if prst == "" {
		prst = "rect"
	}
	return fmt.Sprintf(`      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="%d" name="%s" descr="%s"%s/>
          <p:cNvPicPr>
            <a:picLocks noChangeAspect="1"/>
          </p:cNvPicPr>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          %s
          <a:stretch>
            <a:fillRect/>
          </a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm%s>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
          <a:prstGeom prst="%s">
            <a:avLst/>
          </a:prstGeom>%s
        </p:spPr>
      </p:pic>
`, id, xmlEscape(name), xmlEscape(s.description), hiddenAttr(&s.BaseShape),
		blipFillChildrenXML(relIDNumber(relID), s.alpha, s.cropLeft, s.cropTop, s.cropRight, s.cropBottom, s.lumBright, s.lumContrast, s.hasDuotone, s.duotoneA, s.duotoneB),
		xfrmAttrs(&s.BaseShape),
		s.offsetX, s.offsetY, s.width, s.height,
		prst, shadowXML)
}

// relIDNumber reads the numeric part of an rId string.
func relIDNumber(relID string) int {
	v := 0
	for _, ch := range relID {
		if ch >= '0' && ch <= '9' {
			v = v*10 + int(ch-'0')
		}
	}
	return v
}

// --- Theme ---

func (w *PPTXWriter) writeTheme(zw *zip.Writer) error {
	masters := w.masterPlan
	if len(masters) == 0 {
		masters = []masterPlan{{index: 1}}
	}
	for _, mp := range masters {
		// The source theme part rides back verbatim: the reader stashed its
		// raw bytes so a deck's colour scheme, fonts and fill styles survive
		// the round trip instead of degrading to the built-in Office theme.
		if st := w.presentation.masterStates[mp.masterPath]; st != nil && st.themePath != "" {
			if raw, ok := w.presentation.masterThemeXMLCache[st.themePath]; ok && len(raw) > 0 {
				if err := writeRawXMLToZip(zw, fmt.Sprintf("ppt/theme/theme%d.xml", mp.index), string(raw)); err != nil {
					return err
				}
				continue
			}
		}
		if err := writeRawXMLToZip(zw, fmt.Sprintf("ppt/theme/theme%d.xml", mp.index), defaultThemeXML()); err != nil {
			return err
		}
	}
	return nil
}

// defaultThemeXML is the built-in Office theme the writer emits when the
// source deck carried no theme part the reader could stash.
func defaultThemeXML() string {
	return fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="%s" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>`, nsDrawingML)
}
