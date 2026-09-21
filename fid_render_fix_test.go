package gopresentation

import (
	"bytes"
	"image"
	"image/color"
	"strings"
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
// <a:rPr><a:effectLst><a:outerShdw> — run-level text shadow, and <a:sym>.
// ---------------------------------------------------------------------------

func TestRunShadowRoundTrips(t *testing.T) {
	slide := `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="800000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US" sz="1600" b="1" dirty="0"><a:solidFill><a:srgbClr val="302E27"/></a:solidFill><a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:srgbClr val="000000"><a:alpha val="43137"/></a:srgbClr></a:outerShdw></a:effectLst><a:sym typeface="Symbol"/></a:rPr><a:t>x</a:t></a:r></a:p>
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
	if len(paras) != 1 {
		t.Fatalf("got %d paragraphs, want 1", len(paras))
	}
	var font *Font
	for _, elem := range paras[0].elements {
		if tr, ok := elem.(*TextRun); ok {
			font = tr.font
		}
	}
	if font == nil {
		t.Fatal("no text run found")
	}
	// The reader's units: blur and distance in points, direction in degrees,
	// alpha in percent.
	if font.Shadow == nil || !font.Shadow.Visible {
		t.Fatalf("run shadow not read: %+v", font.Shadow)
	}
	if font.Shadow.BlurRadius != 3 || font.Shadow.Distance != 3 {
		t.Errorf("shadow blur/dist = %d/%d pt, want 3/3", font.Shadow.BlurRadius, font.Shadow.Distance)
	}
	if font.Shadow.Direction != 45 {
		t.Errorf("shadow direction = %d°, want 45 (2700000 in 60000ths)", font.Shadow.Direction)
	}
	if font.Shadow.Alpha != 43 {
		t.Errorf("shadow alpha = %d, want 43 (43137 thousandths)", font.Shadow.Alpha)
	}
	// The reader bakes the alpha into the colour's own byte as well (the shape
	// shadow convention); the renderer takes the real alpha from Shadow.Alpha.
	if !strings.HasSuffix(font.Shadow.Color.ARGB, "000000") {
		t.Errorf("shadow colour = %q, want the black …000000", font.Shadow.Color.ARGB)
	}
	if font.NameSym != "Symbol" {
		t.Errorf("sym typeface = %q, want Symbol — <a:sym> is what carries the run's PUA glyphs", font.NameSym)
	}

	// The writer emits PowerPoint's own values back. The package bytes are a
	// zipped archive, so search inside the unpacked slide part.
	p2 := New()
	shape := p2.GetActiveSlide().CreateRichTextShape()
	tr := shape.CreateTextRun("x")
	f := NewFont()
	f.Size = 16
	f.Shadow = NewShadow()
	f.Shadow.Visible = true
	f.Shadow.BlurRadius = 3
	f.Shadow.Distance = 3
	f.Shadow.Direction = 45
	f.Shadow.Color = NewColor("000000")
	f.Shadow.Alpha = 43
	f.NameSym = "Symbol"
	tr.SetFont(f)
	outParts := zipParts(t, writeToBytes(t, p2))
	slideXML := string(outParts["ppt/slides/slide1.xml"])
	for _, want := range []string{
		`<a:outerShdw blurRad="38100" dist="38100" dir="2700000"`,
		`<a:alpha val="43000"/>`, // the model keeps alpha in percent: 43137‰ rounds to 43%
		`<a:sym typeface="Symbol"/>`,
	} {
		if !bytes.Contains([]byte(slideXML), []byte(want)) {
			t.Errorf("slide part is missing %s:\n%s", want, slideXML)
		}
	}
}

// TestSymbolPUARendersAsTheSymFont pins the rendering half of <a:sym>: a
// U+F000-F0FF character draws from the font the run declares with <a:sym>, not
// as the declared text face's .notdef box. slide25 of the comparison deck
// writes "1 → 2" as U+F0AE with <a:sym typeface="Symbol"/> — the box version
// of that arrow was one of the largest remaining differences from PowerPoint.
func TestSymbolPUARendersAsTheSymFont(t *testing.T) {
	fc := NewFontCache()
	if !fc.HasFont("Symbol", false, false) || !fc.CoversRune("Symbol", false, false, 0xF0AE) {
		t.Skip("no installed Symbol face covers U+F0AE")
	}
	declared := installedFaceLacking(fc, "\uf0ae")
	if declared == "" {
		t.Skip("every installed face draws U+F0AE; the boxed case cannot be constructed")
	}
	// The arrow via <a:sym> and the same character without it must not be the
	// same drawing: without the sym declaration the run falls to the declared
	// face and boxes.
	pres := New()
	shape := pres.GetActiveSlide().CreateRichTextShape()
	shape.BaseShape.SetOffsetX(400000).SetOffsetY(400000)
	shape.BaseShape.SetWidth(6000000).SetHeight(1500000)
	run := shape.CreateTextRun("\uf0ae")
	run.GetFont().SetName(declared).SetSize(40)
	run.GetFont().NameSym = "Symbol"
	opts := DefaultRenderOptions()
	opts.Width = 640
	opts.FontCache = fc
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render with sym: %v", err)
	}
	rect := emuRect(t, pres, 400000, 400000, 6000000, 1500000, 640)
	symInk := countInkIn(img, rect)
	if symInk == 0 {
		t.Fatal("U+F0AE rendered no ink at all with the sym font declared")
	}

	run.GetFont().NameSym = ""
	img2, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render without sym: %v", err)
	}
	plainInk := countInkIn(img2, rect)
	if plainInk == symInk {
		t.Errorf("U+F0AE inked %d pixels with and without <a:sym>: identical counts mean the sym "+
			"declaration was ignored and both drew the declared face's .notdef box", plainInk)
	}
}

// TestTextShadowAddsOffsetInk pins the draw half of the run shadow: the glyph
// pass is repeated offset along the shadow direction before the real text, so
// a shadowed run puts more ink on the canvas than the same run without one.
// The shadow's distance used to be multiplied by a pixels-per-EMU scale
// directly, which rounds any realistic distance to zero — an invisible shadow.
// The blur is deliberately zero: blur spreads ink even at a zero offset, which
// would mask the lost displacement, and the offset alone is the unit bug.
func TestTextShadowAddsOffsetInk(t *testing.T) {
	fc := NewFontCache()
	render := func() int {
		pres := New()
		shape := pres.GetActiveSlide().CreateRichTextShape()
		shape.BaseShape.SetOffsetX(400000).SetOffsetY(400000)
		shape.BaseShape.SetWidth(6000000).SetHeight(1500000)
		run := shape.CreateTextRun("H")
		run.GetFont().SetName("Arial").SetSize(40)
		run.GetFont().Shadow = NewShadow()
		run.GetFont().Shadow.Visible = true
		run.GetFont().Shadow.BlurRadius = 0
		run.GetFont().Shadow.Distance = 6
		run.GetFont().Shadow.Direction = 45
		run.GetFont().Shadow.Color = NewColor("000000")
		run.GetFont().Shadow.Alpha = 100
		opts := DefaultRenderOptions()
		opts.Width = 640
		opts.FontCache = fc
		img, err := pres.SlideToImage(0, opts)
		if err != nil {
			t.Fatalf("render: %v", err)
		}
		return countInkIn(img, emuRect(t, pres, 400000, 400000, 6000000, 1500000, 640))
	}
	plain := func() int {
		pres := New()
		shape := pres.GetActiveSlide().CreateRichTextShape()
		shape.BaseShape.SetOffsetX(400000).SetOffsetY(400000)
		shape.BaseShape.SetWidth(6000000).SetHeight(1500000)
		run := shape.CreateTextRun("H")
		run.GetFont().SetName("Arial").SetSize(40)
		opts := DefaultRenderOptions()
		opts.Width = 640
		opts.FontCache = fc
		img, err := pres.SlideToImage(0, opts)
		if err != nil {
			t.Fatalf("render: %v", err)
		}
		return countInkIn(img, emuRect(t, pres, 400000, 400000, 6000000, 1500000, 640))
	}
	plainInk, shadowedInk := plain(), render()
	if shadowedInk <= plainInk {
		t.Errorf("shadowed run inked %d pixels vs %d plain: the shadow added nothing, so the offset is probably zero", shadowedInk, plainInk)
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

// ---------------------------------------------------------------------------
// Placeholder bodyPr anchoring inherits down the slide → layout → master
// ladder. The master's title placeholder is where anchor="ctr" lives in a
// deck PowerPoint wrote — slide5/6 of the comparison deck drew their titles
// at the top of the frame because nothing read it.
// ---------------------------------------------------------------------------

const anchorMaster = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p/></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst/>
  <p:txStyles/>
</p:sldMaster>`

const anchorSlideTitle = `
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="4" name="Title"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="title"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="300000"/><a:ext cx="8000000" cy="1000000"/></a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US"/><a:t>TITLE</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`

func TestPlaceholderAnchorInheritsFromMaster(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(anchorSlideTitle))
	parts["ppt/slideMasters/slideMaster1.xml"] = []byte(anchorMaster)
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	ph, ok := firstShapeOfType[*PlaceholderShape](pres)
	if !ok {
		t.Fatal("no title placeholder found")
	}
	if ph.textAnchor != TextAnchorMiddle {
		t.Errorf("read anchor = %q, want %q: the master placeholder's anchor=ctr did not come down the ladder", ph.textAnchor, TextAnchorMiddle)
	}
	// The write half: the inherited anchor is emitted on the slide's own
	// bodyPr so the value survives another round trip.
	parts2 := zipParts(t, writeToBytes(t, pres))
	if got := string(parts2["ppt/slides/slide1.xml"]); !bytes.Contains([]byte(got), []byte(`anchor="ctr"`)) {
		t.Errorf("written slide has no anchor=\"ctr\":\n%s", got)
	}
}

// ---------------------------------------------------------------------------
// Master bodyStyle space-before. <a:spcBef><a:spcPct val="20000"/> is a
// percentage of the line — 20% of a 28pt level's 33.6pt line is 6.72pt, the
// gap between every pair of paragraphs on slide32, which used to vanish
// because only spcPts was understood and only for runs, never inherited.
// ---------------------------------------------------------------------------

const spcBefMaster = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    <p:bodyStyle>
      <a:lvl1pPr><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:defRPr sz="2800"/></a:lvl1pPr>
    </p:bodyStyle>
  </p:txStyles>
</p:sldMaster>`

func TestMasterSpaceBeforeReachesParagraphs(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(masterLevelSlideShape))
	parts["ppt/slideMasters/slideMaster1.xml"] = []byte(spcBefMaster)
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
	// 28pt line × 1.2 × 20% = 6.72pt = 672 hundredths.
	if got := paras[0].spaceBefore; got != 672 {
		t.Errorf("level-0 paragraph spaceBefore = %d, want 672 (20%% of a 28pt line)", got)
	}
	// The write half: the baked value survives as absolute points.
	parts2 := zipParts(t, writeToBytes(t, pres))
	if got := string(parts2["ppt/slides/slide1.xml"]); !bytes.Contains([]byte(got), []byte(`<a:spcBef><a:spcPts val="672"/>`)) {
		t.Errorf("written slide has no spcPts val=\"672\":\n%s", got)
	}
}

// A paragraph that declares no bullet inherits the master bodyStyle's bullet
// for its level: slide32's sub-bullets are <a:pPr marL indent> with no bullet
// child at all, and the "•" comes from the master and nowhere else. A
// paragraph that did declare one — <a:buChar>, <a:buAutoNum> or <a:buNone> —
// keeps it, and an empty paragraph draws no glyph even when a bullet would be
// inherited, because PowerPoint renders no bullet for a paragraph with no runs.
const bulletMaster = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    <p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr>
        <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
        <a:buChar char="•"/>
        <a:defRPr sz="3200"/>
      </a:lvl1pPr>
      <a:lvl2pPr>
        <a:buFont typeface="Courier New"/>
        <a:buChar char="»"/>
        <a:defRPr sz="2800"/>
      </a:lvl2pPr>
      <a:lvl3pPr>
        <a:buNone/>
        <a:defRPr sz="2400"/>
      </a:lvl3pPr>
    </p:bodyStyle>
    <p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle>
  </p:txStyles>
</p:sldMaster>`

const bulletSlideShape = `
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
    <a:p><a:r><a:rPr lang="en-US"/><a:t>INHERIT</a:t></a:r></a:p>
    <a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US"/><a:t>INHERIT2</a:t></a:r></a:p>
    <a:p><a:pPr><a:buChar char="-"/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>EXPLICIT</a:t></a:r></a:p>
    <a:p><a:pPr><a:buNone/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>NONE</a:t></a:r></a:p>
    <a:p><a:endParaRPr lang="en-US"/></a:p>
  </p:txBody>
</p:sp>`

func TestMasterBodyStyleBulletIsInherited(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(bulletSlideShape))
	parts["ppt/slideMasters/slideMaster1.xml"] = []byte(bulletMaster)
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
	if len(paras) != 5 {
		t.Fatalf("got %d paragraphs, want 5", len(paras))
	}

	assertChar := func(p *Paragraph, wantChar, wantFont string, label string) {
		t.Helper()
		b := p.GetBullet()
		if b == nil {
			t.Errorf("%s: bullet = nil, want a character bullet drawing %q", label, wantChar)
			return
		}
		if b.Type != BulletTypeChar || b.Style != wantChar {
			t.Errorf("%s: bullet = type %d style %q, want a character bullet drawing %q", label, b.Type, b.Style, wantChar)
		}
		if b.Font != wantFont {
			t.Errorf("%s: bullet font = %q, want %q", label, b.Font, wantFont)
		}
	}
	assertChar(paras[0], "•", "Arial", "level-0 paragraph")
	assertChar(paras[1], "»", "Courier New", "level-1 paragraph")
	assertChar(paras[2], "-", "", "explicit buChar paragraph")
	if b := paras[3].GetBullet(); b != nil && b.Type != BulletTypeNone {
		t.Errorf("buNone paragraph: bullet type = %d, want None", b.Type)
	}
	if b := paras[4].GetBullet(); b != nil && b.Type != BulletTypeNone {
		t.Errorf("empty paragraph: bullet type = %d, want None (no glyph for a runless paragraph)", b.Type)
	}

	// Structural half: what the reader baked in must survive a save, in the
	// element order the schema fixes (buFont before buChar).
	got := string(zipParts(t, writeToBytes(t, pres))["ppt/slides/slide1.xml"])
	for _, want := range []string{`<a:buFont typeface="Arial"/>`, `<a:buChar char="•"/>`, `<a:buChar char="-"/>`, `<a:buNone/>`} {
		if !strings.Contains(got, want) {
			t.Errorf("written slide lost %q:\n%s", want, got)
		}
	}
}

// PowerPoint applies space before the *first* paragraph of a text body too,
// so the renderer must not skip it the way it skips nothing else.
func TestFirstParagraphSpaceBeforeShiftsInk(t *testing.T) {
	fc := NewFontCache()
	firstInkRow := func(spaceBefore int) int {
		pres := New()
		shape := pres.GetActiveSlide().CreateRichTextShape()
		shape.BaseShape.SetOffsetX(400000).SetOffsetY(400000)
		shape.BaseShape.SetWidth(6000000).SetHeight(2000000)
		run := shape.CreateTextRun("H")
		run.GetFont().SetName("Arial").SetSize(24)
		run.GetFont().Color = NewColor("000000")
		shape.GetParagraphs()[0].spaceBefore = spaceBefore
		opts := DefaultRenderOptions()
		opts.Width = 640
		opts.FontCache = fc
		img, err := pres.SlideToImage(0, opts)
		if err != nil {
			t.Fatalf("render: %v", err)
		}
		rect := emuRect(t, pres, 400000, 400000, 6000000, 2000000, 640)
		for y := rect.Min.Y; y < rect.Max.Y; y++ {
			for x := rect.Min.X; x < rect.Max.X; x++ {
				r, g, b, a := img.At(x, y).RGBA()
				if a != 0 && (r+g+b)/3 < 20000 {
					return y
				}
			}
		}
		return rect.Max.Y
	}
	plain := firstInkRow(0)
	spaced := firstInkRow(1000) // 10pt: 8-9px at this scale
	if spaced-plain < 5 {
		t.Errorf("first paragraph ink row moved %d px with spaceBefore=1000 (%d vs %d); the first paragraph's space was skipped", spaced-plain, spaced, plain)
	}
}

// An auto-numbered list keeps counting across the unnumbered sub-paragraphs
// nested under its items: the count is per outline level. slide35 rendered
// "1. 1. 1. 1. 1." because any non-numbered paragraph reset the only counter.
func TestAutoNumberCountsPerLevel(t *testing.T) {
	mk := func(level int, b *Bullet) *Paragraph {
		p := NewParagraph()
		p.alignment = NewAlignment()
		p.alignment.Level = level
		p.bullet = b
		return p
	}
	number := func() *Bullet {
		return &Bullet{Type: BulletTypeAutoNum, StartAt: 1, NumFormat: NumFormatArabicPeriod}
	}
	paras := []*Paragraph{
		mk(0, number()),
		mk(1, &Bullet{Type: BulletTypeNone}),
		mk(0, number()),
		mk(0, number()),
	}
	got := bulletOrdinals(paras)
	want := []int{1, 0, 2, 3}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("ordinals = %v, want %v", got, want)
		}
	}
}

// A tab inside a run advances to the next one-inch tab stop instead of
// drawing the control character — the face has no glyph for it, so the text
// used to grow a .notdef box (slide35's sub-bullets open with a tab, and the
// "Memory:\t224 MB" info tables align on them).
func TestTabJumpsToTheNextStop(t *testing.T) {
	fc := NewFontCache()
	columns := func(text string) []bool {
		pres := New()
		shape := pres.GetActiveSlide().CreateRichTextShape()
		shape.BaseShape.SetOffsetX(200000).SetOffsetY(200000)
		shape.BaseShape.SetWidth(6000000).SetHeight(800000)
		run := shape.CreateTextRun(text)
		run.GetFont().SetName("Arial").SetSize(20)
		run.GetFont().Color = NewColor("000000")
		opts := DefaultRenderOptions()
		opts.Width = 640
		opts.FontCache = fc
		img, err := pres.SlideToImage(0, opts)
		if err != nil {
			t.Fatalf("render %q: %v", text, err)
		}
		rect := emuRect(t, pres, 200000, 200000, 6000000, 800000, 640)
		cols := make([]bool, rect.Dx())
		for y := rect.Min.Y; y < rect.Max.Y; y++ {
			for x := rect.Min.X; x < rect.Max.X; x++ {
				r, g, b, a := img.At(x, y).RGBA()
				if a != 0 && (r+g+b)/3 < 20000 {
					cols[x-rect.Min.X] = true
				}
			}
		}
		return cols
	}
	lastInk := func(cols []bool) int {
		for i := len(cols) - 1; i >= 0; i-- {
			if cols[i] {
				return i
			}
		}
		return -1
	}
	firstInk := func(cols []bool) int {
		for i := range cols {
			if cols[i] {
				return i
			}
		}
		return -1
	}
	plainCols := columns("AB")
	plainEnd := lastInk(plainCols)
	// Plain glyphs sit together: the whole word spans a couple of letters.
	if plainEnd-firstInk(plainCols) > 30 {
		t.Errorf("plain AB spans %d px, want adjacent glyphs", plainEnd-firstInk(plainCols))
	}
	tabEnd := lastInk(columns("A\tB"))
	// One inch is ≥40px at these scales. If the tab jumped, "B" starts a full
	// stop later; if the tab was skipped or drawn as a narrow tofu box, "B"
	// sits within a glyph width of "A".
	if tabEnd-plainEnd < 30 {
		t.Errorf("tabbed text's last ink column = %d vs %d for plain AB; the tab advanced %d px, want a full stop (~≥40px) — a skipped or tofu-drawn tab advances nothing",
			tabEnd, plainEnd, tabEnd-plainEnd)
	}
}

// PowerPoint draws the bullet at marL+indent but starts the text at marL —
// the bullet is followed by something tab-like, not by its own advance. The
// renderer used to flow the text right after the bullet glyph, so every
// bulleted line sat a bullet-width too far left (slide34's "–" lists).
//
// The reference is the same paragraph without a bullet: its text starts on
// marL by construction, so the bulleted render must put its text ink on the
// same columns regardless of where the bullet itself lands.
func TestBulletTextLandsOnTheMargin(t *testing.T) {
	fc := NewFontCache()
	// The text face (Calibri) is deliberately shorter than the bullet face
	// (Arial): PowerPoint sizes the line by its text, so a tall bullet face
	// must not push the line's baseline down.
	build := func(bulleted bool) [][]bool {
		pres := New()
		shape := pres.GetActiveSlide().CreateRichTextShape()
		shape.BaseShape.SetOffsetX(200000).SetOffsetY(200000)
		shape.BaseShape.SetWidth(6000000).SetHeight(800000)
		run := shape.CreateTextRun("Wed")
		run.GetFont().SetName("Calibri").SetSize(20)
		run.GetFont().Color = NewColor("000000")
		para := shape.GetParagraphs()[0]
		if bulleted {
			para.bullet = &Bullet{Type: BulletTypeChar, Style: "•", Font: "Arial"}
		}
		// One inch of margin, half an inch of hanging indent: 64px and 32px
		// at this render scale. The bullet pen sits at inset+32px; the text
		// must land on inset+64px either way.
		para.alignment = NewAlignment()
		para.alignment.MarginLeft = 914400
		para.alignment.Indent = -457200

		opts := DefaultRenderOptions()
		opts.Width = 640
		opts.FontCache = fc
		img, err := pres.SlideToImage(0, opts)
		if err != nil {
			t.Fatalf("render: %v", err)
		}
		grid := make([][]bool, 480)
		for y := range grid {
			grid[y] = make([]bool, 640)
		}
		for y := 0; y < 480; y++ {
			for x := 0; x < 640; x++ {
				r, g, b, a := img.At(x, y).RGBA()
				if a != 0 && (r+g+b)/3 < 20000 {
					grid[y][x] = true
				}
			}
		}
		return grid
	}
	firstInk := func(grid [][]bool) int {
		for x := range grid[0] {
			for y := range grid {
				if grid[y][x] {
					return x
				}
			}
		}
		return -1
	}
	// First ink row strictly to the right of column from: the text glyph's
	// top edge, unaffected by the bullet's ink.
	firstRowFrom := func(grid [][]bool, from int) int {
		for y := range grid {
			for x := from; x < len(grid[y]); x++ {
				if grid[y][x] {
					return y
				}
			}
		}
		return -1
	}
	plainGrid := build(false)
	plainAt := firstInk(plainGrid)
	plainRow := firstRowFrom(plainGrid, plainAt)

	// The bulleted render carries the bullet's own ink first; find where the
	// text begins by skipping the leading gap between the bullet and it.
	bulletedGrid := build(true)
	bulletedCols := make([]bool, len(bulletedGrid[0]))
	for y := range bulletedGrid {
		for x, on := range bulletedGrid[y] {
			if on {
				bulletedCols[x] = true
			}
		}
	}
	gap := -1
	sawInk := false
	var textAt int
	for i, on := range bulletedCols {
		if on {
			if !sawInk {
				sawInk = true
			} else if gap >= 0 {
				textAt = i
				break
			}
			gap = -1
		} else if sawInk {
			if gap < 0 {
				gap = 0
			}
			gap++
		}
	}
	if textAt <= 0 {
		t.Fatal("no second ink interval found after the bullet")
	}
	if d := textAt - plainAt; d < 29 || d > 35 {
		t.Errorf("bulleted text starts %d px after the plain paragraph's first line (%d vs %d), want one hang = 32 px; the text is riding the bullet's own advance instead of starting on marL", d, textAt, plainAt)
	}
	// Vertical: the taller bullet face must not grow the line box. With the
	// bullet's metrics counted, the text's baseline (and its ink) slides
	// down by the ascent difference.
	textRow := firstRowFrom(bulletedGrid, textAt)
	if d := textRow - plainRow; d < -1 || d > 1 {
		t.Errorf("bulleted text ink row = %d vs %d plain; the bullet face's metrics grew the line box and pushed the text down %d px", textRow, plainRow, d)
	}
}

// A shape marked <p:cNvPr hidden="1"> must stay in the file but never reach
// the canvas. The real deck's slide27 hides a white-filled rectangle drawn on
// top of an entire diagram; the renderer used to draw it and blank the region.
// Asserted on the writer (the attribute is emitted), the reader (it round
// trips), and the renderer (the ink of an earlier shape survives a hidden
// white cover drawn later in z-order).
func TestHiddenShapeIsKeptButNotDrawn(t *testing.T) {
	pres := New()
	sl := pres.GetActiveSlide()

	text := sl.CreateRichTextShape()
	text.BaseShape.SetOffsetX(300000).SetOffsetY(300000)
	text.BaseShape.SetWidth(3000000).SetHeight(800000)
	run := text.CreateTextRun("INK")
	run.GetFont().SetName("Arial").SetSize(24)
	run.GetFont().Color = NewColor("000000")

	cover := NewAutoShape()
	cover.BaseShape.SetOffsetX(200000).SetOffsetY(200000)
	cover.BaseShape.SetWidth(3400000).SetHeight(1200000)
	cover.SetFill(NewFill().SetSolid(NewColor("FFFFFF")))
	cover.SetHidden(true)
	sl.AddShape(cover) // later in z-order: would cover the text if drawn

	// Write half: the hidden attribute is on the emitted cNvPr, and a shape
	// never told to hide emits nothing.
	parts := zipParts(t, writeToBytes(t, pres))
	slideXML := string(parts["ppt/slides/slide1.xml"])
	if !strings.Contains(slideXML, `hidden="1"`) {
		t.Errorf("written slide does not mark the cover shape hidden:\n%s", slideXML)
	}
	if strings.Count(slideXML, `hidden="1"`) != 1 {
		t.Errorf("hidden attribute emitted %d times, want exactly 1 (only the cover)", strings.Count(slideXML, `hidden="1"`))
	}

	// Read half: the flag survives the round trip on the right shape.
	data := writeToBytes(t, pres)
	pres2, err := ReadFrom(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("re-open: %v", err)
	}
	shapes := pres2.GetActiveSlide().GetShapes()
	if len(shapes) != 2 {
		t.Fatalf("got %d shapes after round trip, want 2", len(shapes))
	}
	if shapes[0].base().hidden {
		t.Errorf("text shape came back hidden; only the cover was marked")
	}
	if !shapes[1].base().hidden {
		t.Errorf("cover shape lost its hidden flag on the round trip")
	}

	// Render half: the text's ink must survive the hidden white cover above it.
	fc := NewFontCache()
	opts := DefaultRenderOptions()
	opts.Width = 640
	opts.FontCache = fc
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	rect := emuRect(t, pres, 300000, 300000, 3000000, 800000, 640)
	dark := 0
	for y := rect.Min.Y; y < rect.Max.Y; y++ {
		for x := rect.Min.X; x < rect.Max.X; x++ {
			r, g, b, a := img.At(x, y).RGBA()
			if a != 0 && (r+g+b)/3 < 20000 {
				dark++
			}
		}
	}
	if dark < 50 {
		t.Errorf("text region has %d dark px; the hidden white cover is being drawn over the text", dark)
	}
}

// A paragraph that states marL="0" indent="0" is overriding the master's
// hanging indent, not leaving it unset — 0 and "absent" are different claims.
// The inheritance ladder used to see MarginLeft == 0 and bake the master
// bodyStyle's marL in, shifting every such paragraph right by the margin
// (60px on the fidelity deck's slide27). The writer must re-emit the explicit
// zero, or the override is lost on the next save.
const zeroMarginMaster = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    <p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr marL="342900" indent="-342900"><a:defRPr sz="3200"/></a:lvl1pPr>
    </p:bodyStyle>
    <p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle>
  </p:txStyles>
</p:sldMaster>`

const zeroMarginSlideShape = `
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
    <a:p><a:pPr marL="0" indent="0"/><a:r><a:rPr lang="en-US"/><a:t>ZERO</a:t></a:r></a:p>
    <a:p><a:r><a:rPr lang="en-US"/><a:t>INHERIT</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`

func TestExplicitZeroMarginIsNotOverridden(t *testing.T) {
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(zeroMarginSlideShape))
	parts["ppt/slideMasters/slideMaster1.xml"] = []byte(zeroMarginMaster)
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
	if paras[0].alignment == nil {
		t.Fatal("explicit-zero paragraph has no alignment")
	}
	if paras[0].alignment.MarginLeft != 0 || paras[0].alignment.Indent != 0 {
		t.Errorf("explicit-zero paragraph came back marL=%d indent=%d; the master's hanging indent overrode the slide's marL=\"0\" indent=\"0\"",
			paras[0].alignment.MarginLeft, paras[0].alignment.Indent)
	}
	if paras[1].alignment == nil || paras[1].alignment.MarginLeft != 342900 || paras[1].alignment.Indent != -342900 {
		gotML, gotInd := int64(-1), int64(-1)
		if paras[1].alignment != nil {
			gotML, gotInd = paras[1].alignment.MarginLeft, paras[1].alignment.Indent
		}
		t.Errorf("unset paragraph should inherit the master's marL=342900 indent=-342900, got marL=%d indent=%d", gotML, gotInd)
	}

	// The explicit zero must survive a save as a stated attribute.
	got := string(zipParts(t, writeToBytes(t, pres))["ppt/slides/slide1.xml"])
	if !strings.Contains(got, `<a:pPr marL="0" indent="0">`) {
		t.Errorf("written slide dropped the explicit marL=\"0\" indent=\"0\":\n%s", got)
	}
}
