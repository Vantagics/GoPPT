package gopresentation

import (
	"archive/zip"
	"fmt"
	"strings"
)

// --- Presentation Part ---

func (w *PPTXWriter) writePresentation(zw *zip.Writer) error {
	layout := w.presentation.layout

	slideList := ""
	relIdx := 2 // rId1 is slideMaster
	for i := range w.presentation.slides {
		slideList += fmt.Sprintf(`    <p:sldId id="%d" r:id="rId%d"/>
`, 256+i, relIdx)
		relIdx++
	}

	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
%s  </p:sldIdLst>
  <p:sldSz cx="%d" cy="%d" type="%s"/>
  <p:notesSz cx="%d" cy="%d"/>
  <p:defaultTextStyle/>
</p:presentation>`,
		nsDrawingML, nsOfficeDocRels, nsPresentationML,
		slideList,
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
	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">
  <p:cSld>
    <p:bg>
      <p:bgRef idx="1001">
        <a:schemeClr val="bg1"/>
      </p:bgRef>
    </p:bg>
    <p:spTree>
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
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>`, nsDrawingML, nsOfficeDocRels, nsPresentationML)

	if err := writeRawXMLToZip(zw, "ppt/slideMasters/slideMaster1.xml", content); err != nil {
		return err
	}

	// Slide master rels
	rels := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="%s">
  <Relationship Id="rId1" Type="%s" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="%s" Target="../theme/theme1.xml"/>
</Relationships>`, nsRelationships, relTypeSlideLayout, relTypeTheme)
	return writeRawXMLToZip(zw, "ppt/slideMasters/_rels/slideMaster1.xml.rels", rels)
}

// --- Slide Layout ---

func (w *PPTXWriter) writeSlideLayout(zw *zip.Writer) error {
	return w.writeSlideLayoutN(zw, w.layoutPlan[0])
}

// writeSlideLayoutN emits one layout part with its decoration shapes — the
// layout deco is where decks like KEY02 keep their whole visual identity —
// and writes the deco pictures' media alongside, wiring the layout's own
// rels to them.
func (w *PPTXWriter) writeSlideLayoutN(zw *zip.Writer, entry layoutPlanEntry) error {
	decoXML := strings.Builder{}
	mediaRels := ""
	shapeID := 2
	if len(entry.decoShapes) > 0 {
		mediaIdx := 0
		var body strings.Builder
		for _, sh := range entry.decoShapes {
			switch s := sh.(type) {
			case *DrawingShape:
				if s.data != nil {
					// A deco picture: its media is written beside the layout
					// part and referenced from the layout's own rels.
					mediaIdx++
					mediaName := fmt.Sprintf("imageL%d_%d.%s", entry.index, mediaIdx, w.getImageExtension(s))
					fw, ferr := zw.Create("ppt/media/" + mediaName)
					if ferr != nil {
						return ferr
					}
					if _, werr := fw.Write(s.data); werr != nil {
						return werr
					}
					relID := fmt.Sprintf("rId%d", 1+mediaIdx)
					mediaRels += fmt.Sprintf("  <Relationship Id=\"%s\" Type=\"%s\" Target=\"../media/%s\"/>\n", relID, relTypeImage, mediaName)


					continue
				}
				decoXML.WriteString(w.writeDrawingShapeXML(s, &shapeID, 1))
			case *AutoShape:
				decoXML.WriteString(w.writeAutoShapeXML(s, &shapeID))
			case *RichTextShape:
				decoXML.WriteString(w.writeRichTextShapeXML(s, &shapeID))
			case *LineShape:
				decoXML.WriteString(w.writeLineShapeXML(s, &shapeID))
			case *GroupShape:
				decoXML.WriteString(w.writeGroupShapeXML(s, &shapeID, 1))
			case *PlaceholderShape:
				// A placeholder in the layout deco: its fill/border paints.
				decoXML.WriteString(w.writeRichTextShapeXML(&s.RichTextShape, &shapeID))
			}
		}
		decoXML.WriteString(body.String())
	}

	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="%s" xmlns:r="%s" xmlns:p="%s" type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
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
</p:sldLayout>`, nsDrawingML, nsOfficeDocRels, nsPresentationML, decoXML.String())

	if err := writeRawXMLToZip(zw, fmt.Sprintf("ppt/slideLayouts/slideLayout%d.xml", entry.index), content); err != nil {
		return err
	}

	// Slide layout rels: the master, then the deco pictures' media.
	rels := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="%s">
  <Relationship Id="rId1" Type="%s" Target="../slideMasters/slideMaster1.xml"/>
%s</Relationships>`, nsRelationships, relTypeSlideMaster, mediaRels)
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
	content := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
	return writeRawXMLToZip(zw, "ppt/theme/theme1.xml", content)
}
