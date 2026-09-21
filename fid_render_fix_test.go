package gopresentation

import (
	"bytes"
	"image"
	"image/color"
	"testing"
)

// Regression tests for the fidelity fixes that came out of the
// PowerPoint-COM pixel comparison campaign (out_deck/cmp). Each test pins
// one fix to the function that carries it, so a future break fails here
// with a message that names the feature instead of only moving a
// whole-deck diff percentage.

// ---------------------------------------------------------------------------
// applyColorTransforms — the HLS folding behind table style band fills.
// ---------------------------------------------------------------------------

// Office 2007's "accent1, lighter 40%" is tint 40000, and PowerPoint renders
// it as #95B3D7. The transform works in HLS luminance, so a wrong scale (an
// 0-255 value fed where 0-1 was expected, the bug this test guards) collapses
// everything toward black.
func TestApplyColorTransformsTint40MatchesOffice(t *testing.T) {
	got := applyColorTransforms(NewColor("4F81BD"), 0.4, -1, -1, -1)
	if want := (Color{ARGB: "FF95B3D7"}); got != want {
		t.Errorf("tint 40%% of 4F81BD = %s, want FF95B3D7", got.ARGB)
	}
}

// <a:tint val="40000"/> and <a:lumMod val="60000"/><a:lumOff val="40000"/>
// are two spellings of the same PowerPoint command, so the two code paths
// must land on the same colour.
func TestApplyColorTransformsTintEqualsLumModLumOff(t *testing.T) {
	viaTint := applyColorTransforms(NewColor("4F81BD"), 0.4, -1, -1, -1)
	viaLum := applyColorTransforms(NewColor("4F81BD"), -1, -1, 0.6, 0.4)
	if viaTint != viaLum {
		t.Errorf("tint path %s != lumMod+lumOff path %s", viaTint.ARGB, viaLum.ARGB)
	}
}

// ---------------------------------------------------------------------------
// parseTableStyles — the reader half of ppt/tableStyles.xml.
// ---------------------------------------------------------------------------

const tableStylesFixture = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">
  <a:tblStyle styleId="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">
    <a:wholeTbl>
      <a:tcPr>
        <a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill>
      </a:tcPr>
    </a:wholeTbl>
    <a:band1H>
      <a:tcPr>
        <a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill>
      </a:tcPr>
    </a:band1H>
    <a:band2H>
      <a:tcPr>
        <a:fill><a:solidFill><a:srgbClr val="E9EDF4"/></a:solidFill></a:fill>
      </a:tcPr>
    </a:band2H>
    <a:band1V>
      <a:tcPr>
        <a:tcBdr><a:lnL w="12700"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:lnL></a:tcBdr>
      </a:tcPr>
    </a:band1V>
    <a:firstRow>
      <a:tcTxStyle b="1"/>
    </a:firstRow>
  </a:tblStyle>
</a:tblStyleLst>`

func TestParseTableStylesReadsBandsAndDefault(t *testing.T) {
	styles, def := parseTableStyles([]byte(tableStylesFixture))
	if def != "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" {
		t.Errorf("default style id = %q", def)
	}
	st := styles[def]
	if st == nil {
		t.Fatal("the default style was not parsed")
	}
	if !st.wholeTbl.hasFill || st.wholeTbl.scheme != "accent1" || st.wholeTbl.tint != 0.2 {
		t.Errorf("wholeTbl band = %+v, want accent1 tint 0.2 with a fill", st.wholeTbl)
	}
	if !st.band1H.hasFill || st.band1H.scheme != "accent1" || st.band1H.tint != 0.4 {
		t.Errorf("band1H = %+v, want accent1 tint 0.4 with a fill", st.band1H)
	}
	// A literal colour sits in the fill without a schemeClr wrapper; before
	// the fix the srgbClr branch demanded inScheme and dropped it.
	if !st.band2H.hasFill || st.band2H.scheme != "srgb:E9EDF4" {
		t.Errorf("band2H = %+v, want the literal srgb:E9EDF4 fill", st.band2H)
	}
	// A tcBdr line colour is a schemeClr too, but it is not the band's fill.
	if st.band1V.hasFill {
		t.Errorf("band1V picked up its border line colour as a fill: %+v", st.band1V)
	}
	if !st.firstRow.bold {
		t.Errorf("firstRow bold flag lost: %+v", st.firstRow)
	}
}

// ---------------------------------------------------------------------------
// applyColorReplaces — the <a:clrChange> recolouring on pictures.
// ---------------------------------------------------------------------------

func TestApplyColorReplacesMakesMatchedColorTransparent(t *testing.T) {
	src := image.NewNRGBA(image.Rect(0, 0, 3, 1))
	src.Set(0, 0, color.NRGBA{R: 0, G: 0, B: 0, A: 255})    // matches rule 1
	src.Set(1, 0, color.NRGBA{R: 0, G: 0, B: 1, A: 255})    // matches rule 2
	src.Set(2, 0, color.NRGBA{R: 10, G: 20, B: 30, A: 255}) // matches nothing

	rules := []colorReplace{
		{From: "000000", To: "FFFFFF", ToAlpha: 0}, // black -> transparent white
		{From: "000001", To: "00FF00", ToAlpha: -1},
	}
	out := applyColorReplaces(src, rules)

	r0c := out.At(0, 0).(color.NRGBA)
	if r0c.A != 0 {
		t.Errorf("matched pixel: alpha = %d, want 0 (the colour was keyed out)", r0c.A)
	}
	if r0c.R != 255 || r0c.G != 255 || r0c.B != 255 {
		t.Errorf("matched pixel: rgb = %d,%d,%d, want white", r0c.R, r0c.G, r0c.B)
	}
	// Rule 2 keeps the source alpha (ToAlpha -1), only the colour changes.
	if c1 := out.At(1, 0).(color.NRGBA); c1.A != 255 || c1.G != 255 {
		t.Errorf("rule with ToAlpha -1: got {%d %d %d %d}, want green at full alpha", c1.R, c1.G, c1.B, c1.A)
	}
	// Unmatched pixels pass through untouched.
	wantR, wantG, wantB, wantA := src.At(2, 0).RGBA()
	if gotR, gotG, gotB, gotA := out.At(2, 0).RGBA(); gotR != wantR || gotG != wantG || gotB != wantB || gotA != wantA {
		t.Errorf("unmatched pixel changed: {%d %d %d %d} -> {%d %d %d %d}", wantR, wantG, wantB, wantA, gotR, gotG, gotB, gotA)
	}
}

// ---------------------------------------------------------------------------
// <p:style> fillRef/lnRef — the theme fallback when spPr declares nothing.
// ---------------------------------------------------------------------------

const styleRefSlide = `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Themed"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
  <p:style>
    <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
  </p:style>
</p:sp>
<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="Explicit"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="3000000" y="500000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
  <p:style>
    <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
  </p:style>
</p:sp>
<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="4" name="Arrow"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="2000000"/><a:ext cx="1500000" cy="200000"/></a:xfrm>
    <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
    <a:ln w="19050"/>
  </p:spPr>
  <p:style>
    <a:lnRef idx="2"><a:schemeClr val="accent2"/></a:lnRef>
  </p:style>
</p:cxnSp>`

func TestStyleRefFallsBackToThemeColours(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(styleRefSlide))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}

	var themed, explicit *RichTextShape
	var arrow *LineShape
	for _, sh := range pres.GetAllSlides()[0].GetShapes() {
		switch s := sh.(type) {
		case *RichTextShape:
			if themed == nil {
				themed = s
			} else {
				explicit = s
			}
		case *LineShape:
			arrow = s
		}
	}
	if themed == nil || explicit == nil || arrow == nil {
		t.Fatalf("fixture shapes missing: themed=%v explicit=%v arrow=%v", themed != nil, explicit != nil, arrow != nil)
	}

	// fillRef with no competing spPr fill: the theme accent becomes the fill.
	if themed.fill == nil || themed.fill.Type != FillSolid {
		t.Fatalf("themed shape has no solid fill: %+v", themed.fill)
	}
	if themed.fill.Color != (Color{ARGB: "FF4472C4"}) {
		t.Errorf("fillRef colour = %s, want FF4472C4 (accent1)", themed.fill.Color.ARGB)
	}

	// The style reference is only a fallback: an explicit spPr fill wins.
	if explicit.fill == nil || explicit.fill.Color != (Color{ARGB: "FFFF0000"}) {
		t.Errorf("explicit fill overwritten by the style ref: %+v", explicit.fill)
	}

	// A connector whose <a:ln> names only a width takes its colour from lnRef.
	if arrow.lineColor != (Color{ARGB: "FFED7D31"}) {
		t.Errorf("connector lnRef colour = %s, want FFED7D31 (accent2)", arrow.lineColor.ARGB)
	}
}

// ---------------------------------------------------------------------------
// Master <p:bodyStyle> levels — lvl is 0-based on the paragraph, lvlNpPr is
// 1-based on the master.
// ---------------------------------------------------------------------------

const masterLevelSlideShape = `
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="4" name="Body Placeholder"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="4000000" cy="3000000"/></a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US"/><a:t>LEVEL0</a:t></a:r></a:p>
    <a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US"/><a:t>LEVEL1</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`

const masterLevelMaster = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst/>
  <p:txStyles>
    <p:titleStyle>
      <a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr>
    </p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr>
      <a:lvl2pPr><a:defRPr sz="2800"/></a:lvl2pPr>
    </p:bodyStyle>
    <p:otherStyle>
      <a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr>
    </p:otherStyle>
  </p:txStyles>
</p:sldMaster>`

func TestMasterBodyStyleLevelIsOneBased(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(masterLevelSlideShape))
	parts["ppt/slideMasters/slideMaster1.xml"] = []byte(masterLevelMaster)
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}

	ph, ok := firstShapeOfType[*PlaceholderShape](pres)
	if !ok {
		t.Fatal("no body placeholder found")
	}
	paras := ph.GetParagraphs()
	if len(paras) != 2 {
		t.Fatalf("got %d paragraphs, want 2", len(paras))
	}
	runSize := func(p *Paragraph) int {
		for _, elem := range p.elements {
			if tr, ok := elem.(*TextRun); ok && tr.font != nil {
				return tr.font.Size
			}
		}
		return -1
	}
	// Level 0 -> lvl1pPr (44pt); level 1 -> lvl2pPr (28pt). The off-by-one
	// read lvl2pPr for level 0 and lvl1pPr's 44pt for level 1.
	if got := runSize(paras[0]); got != 44 {
		t.Errorf("level-0 paragraph size = %v, want 44 (lvl1pPr)", got)
	}
	if got := runSize(paras[1]); got != 28 {
		t.Errorf("level-1 paragraph size = %v, want 28 (lvl2pPr); the level ladder is misaligned", got)
	}
}

// ---------------------------------------------------------------------------
// <a:fontRef> must not poison the fillRef colour.
// ---------------------------------------------------------------------------

// A styled shape's <p:style> names four references. The fontRef's schemeClr
// used to be captured as the fill reference's colour, because inFillRef was
// never reset when fillRef closed — a fontRef naming lt1 (white, the usual
// "minor" text colour) turned every fillRef fill white. slide27's database
// cylinders rendered as white boxes in round 16 because of exactly this.
const poisonedStyleShape = `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Disk"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="2000000" cy="1500000"/></a:xfrm>
    <a:prstGeom prst="flowChartMagneticDisk"><a:avLst/></a:prstGeom>
    <a:ln><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p/></p:txBody>
  <p:style>
    <a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>
    <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
    <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
  </p:style>
</p:sp>`

func TestStyleRefFontRefDoesNotPoisonFillRef(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(poisonedStyleShape))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	sh, ok := firstShapeOfType[*AutoShape](pres)
	if !ok {
		t.Fatal("no AutoShape found")
	}
	if sh.fill == nil || sh.fill.Type != FillSolid || sh.fill.Color != (Color{ARGB: "FF4472C4"}) {
		t.Errorf("fillRef colour = %+v, want solid FF4472C4; the fontRef colour leaked into it", sh.fill)
	}
	// The explicit <a:ln> colour still wins over the lnRef.
	if sh.border == nil || sh.border.Color != (Color{ARGB: "FF44546A"}) {
		t.Errorf("line colour = %+v, want the explicit tx2 (FF44546A)", sh.border)
	}
}

// ---------------------------------------------------------------------------
// <a:rPr baseline="…"> — superscript and subscript.
// ---------------------------------------------------------------------------

func TestRunBaselineRoundTrips(t *testing.T) {
	slide := `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="800000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US" sz="1800" baseline="30000"/><a:t>SUP</a:t></a:r></a:p>
    <a:p><a:r><a:rPr lang="en-US" sz="1800" baseline="-25000"/><a:t>SUB</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(slide))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	sh, ok := firstShapeOfType[*RichTextShape](pres)
	if !ok {
		t.Fatal("no RichTextShape found")
	}
	paras := sh.GetParagraphs()
	if len(paras) != 2 {
		t.Fatalf("got %d paragraphs, want 2", len(paras))
	}
	runFont := func(p *Paragraph) *Font {
		for _, elem := range p.elements {
			if tr, ok := elem.(*TextRun); ok {
				return tr.font
			}
		}
		return nil
	}
	sup, sub := runFont(paras[0]), runFont(paras[1])
	if sup == nil || !sup.Superscript || sup.Subscript {
		t.Errorf("baseline=\"30000\" read as Superscript=%v Subscript=%v, want true/false", sup.Superscript, sup.Subscript)
	}
	if sub == nil || !sub.Subscript || sub.Superscript {
		t.Errorf("baseline=\"-25000\" read as Superscript=%v Subscript=%v, want false/true", sub.Superscript, sub.Subscript)
	}

	// The writer emits PowerPoint's own two values. The package bytes are a
	// zipped archive, so search inside the unpacked slide part, not the raw bytes.
	p2 := New()
	shape := p2.GetActiveSlide().CreateRichTextShape()
	tr := shape.CreateTextRun("x")
	f := NewFont()
	f.Size = 18
	f.Superscript = true
	tr.SetFont(f)
	outParts := zipParts(t, writeToBytes(t, p2))
	slideXML := string(outParts["ppt/slides/slide1.xml"])
	if !bytes.Contains([]byte(slideXML), []byte(`baseline="30000"`)) {
		t.Errorf("superscript run wrote no baseline=\"30000\":\n%s", slideXML)
	}
	p3 := New()
	shape = p3.GetActiveSlide().CreateRichTextShape()
	tr = shape.CreateTextRun("x")
	f = NewFont()
	f.Size = 18
	f.Subscript = true
	tr.SetFont(f)
	outParts = zipParts(t, writeToBytes(t, p3))
	slideXML = string(outParts["ppt/slides/slide1.xml"])
	if !bytes.Contains([]byte(slideXML), []byte(`baseline="-25000"`)) {
		t.Errorf("subscript run wrote no baseline=\"-25000\":\n%s", slideXML)
	}
}

// ---------------------------------------------------------------------------
// can / flowChartMagneticDisk — the cylinder preset geometry.
// ---------------------------------------------------------------------------

func TestCylinderPresetsRenderFilled(t *testing.T) {
	fc := NewFontCache()
	for _, prst := range []AutoShapeType{AutoShapeCan, AutoShapeFlowChartMagneticDisk} {
		p := New()
		sh := NewAutoShape()
		sh.shapeType = prst
		sh.SetOffsetX(1000000)
		sh.SetOffsetY(1000000)
		sh.SetWidth(2000000)
		sh.SetHeight(2000000)
		sh.fill = NewFill()
		sh.fill.SetSolid(NewColor("4472C4"))
		p.GetActiveSlide().AddShape(sh)
		img, err := p.SlideToImage(0, goldenOptions(fc))
		if err != nil {
			t.Fatalf("%s: render: %v", prst, err)
		}
		layout := p.GetLayout()
		scale := float64(img.Bounds().Dx()) / float64(layout.CX)
		// The body's centre, below the top ellipse: 60% down the shape.
		cx := int(float64(1000000+2000000/2) * scale)
		cy := int(float64(1000000+2000000*6/10) * scale)
		r, g, bl, a := img.At(cx, cy).RGBA()
		if a == 0 || (r > 0xC000 && g > 0xC000 && bl > 0xC000) {
			t.Errorf("%s: body centre (%d,%d) is empty/white {%d %d %d %d}; the cylinder did not fill",
				prst, cx, cy, r>>8, g>>8, bl>>8, a>>8)
		}
	}
}
