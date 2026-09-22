package gopresentation

import (
	"bytes"
	"strings"
	"testing"
)

// The <p:style> fillRef/lnRef tests.
//
// A style reference does not name a colour — it names an entry of the theme's
// fillStyleLst/lnStyleLst, whose bodies are built from phClr (whatever colour
// the reference carries) passed through transforms. Resolving idx=2 as a solid
// of the scheme colour painted theme-gradient text boxes flat black, and a
// fixed 1pt line weight drew every themed border at the wrong width.

// themeTestTheme is a minimal theme1.xml: the Office palette plus a fmtScheme
// whose second fill style is the standard subtle gradient (tint 50% → tint
// 15%, running bottom-to-top) and whose second line style is 2pt plain.
func themeTestTheme(fillStyleLst, lnStyleLst string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="` + nsDrawingML + `" name="Test">
  <a:themeElements>
    <a:clrScheme name="Test">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Test">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Test">
      <a:fillStyleLst>` + fillStyleLst + `</a:fillStyleLst>
      <a:lnStyleLst>` + lnStyleLst + `</a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`
}

const themeOfficeFillStyles = `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
	`<a:gradFill rotWithShape="1"><a:gsLst>` +
	`<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs>` +
	`<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/></a:schemeClr></a:gs>` +
	`</a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>`

const themeOfficeLnStyles = `<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>` +
	`<a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>`

// themeTestRead reads a one-slide package with the given theme and shape.
func themeTestRead(t *testing.T, theme, shape, slideRels string) *Presentation {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(shape))
	parts["ppt/theme/theme1.xml"] = []byte(theme)
	if slideRels != "" {
		parts["ppt/slides/_rels/slide1.xml.rels"] = []byte(slideRels)
	}
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	return pres
}

// styledShape is a shape whose spPr declares no fill and no line: everything
// it draws comes from the style reference.
func styledShape(fillRef, lnRef string) string {
	return `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Styled"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:style>` + lnRef + fillRef + `
    <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
  </p:style>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`
}

// TestFillRefIdx2IsTheThemeGradient covers the half the solid reading got
// wrong: fillRef idx="2" with phClr=dk1 must render the fillStyleLst's second
// body — a bottom-to-top gradient of the reference colour tinted 50% → 15%,
// i.e. grey 188 under grey 237, not solid black (slide07 drew #000 where
// PowerPoint drew exactly this ramp).
func TestFillRefIdx2IsTheThemeGradient(t *testing.T) {
	theme := themeTestTheme(themeOfficeFillStyles, themeOfficeLnStyles)
	pres := themeTestRead(t, theme, styledShape(
		`<a:fillRef idx="2"><a:schemeClr val="dk1"/></a:fillRef>`, ""), "")
	sh := pres.GetAllSlides()[0].GetShapes()
	if len(sh) == 0 {
		t.Fatal("the slide has no shapes")
	}
	rt, ok := sh[0].(*RichTextShape)
	if !ok {
		t.Fatalf("shape is %T, want *RichTextShape", sh[0])
	}
	if rt.fill == nil || rt.fill.Type != FillGradientLinear {
		t.Fatalf("fillRef idx=2 resolved to %+v, want a linear gradient", rt.fill)
	}
	start := NewColor("FF000000")
	applyTint(&start, 0.5)
	if rt.fill.Color != start {
		t.Errorf("gradient start = %s, want %s (dk1 tinted 50%%)", rt.fill.Color.ARGB, start.ARGB)
	}
	end := NewColor("FF000000")
	applyTint(&end, 0.15)
	if rt.fill.EndColor != end {
		t.Errorf("gradient end = %s, want %s (dk1 tinted 15%%)", rt.fill.EndColor.ARGB, end.ARGB)
	}
	if rt.fill.Rotation != 270 {
		t.Errorf("gradient angle = %d, want 270 (the style's ang=16200000)", rt.fill.Rotation)
	}
}

// TestLnRefTakesThemeWidthAndShade: the line a lnRef draws is the lnStyleLst
// body with the reference colour substituted — including the width that body
// declares (idx=2 is 2pt here, not the 1pt the reader used to hard-code) and
// the transforms the reference's own schemeClr carries (shade 50% turns
// accent1 #4F81BD into #385D8A, the value the COM export shows).
func TestLnRefTakesThemeWidthAndShade(t *testing.T) {
	theme := themeTestTheme(themeOfficeFillStyles, themeOfficeLnStyles)
	pres := themeTestRead(t, theme, styledShape("",
		`<a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>`), "")
	sh := pres.GetAllSlides()[0].GetShapes()
	if len(sh) == 0 {
		t.Fatal("the slide has no shapes")
	}
	rt, ok := sh[0].(*RichTextShape)
	if !ok {
		t.Fatalf("shape is %T, want *RichTextShape", sh[0])
	}
	if rt.border == nil {
		t.Fatal("the lnRef produced no border")
	}
	if rt.border.Width != 2 {
		t.Errorf("border width = %dpt, want 2 (the lnStyleLst idx=2 body's w)", rt.border.Width)
	}
	want := NewColor("FF4F81BD")
	applyShade(&want, 0.5)
	if rt.border.Color != want {
		t.Errorf("border colour = %s, want %s (accent1 shaded 50%%)", rt.border.Color.ARGB, want.ARGB)
	}
}

// TestFillRefWithoutThemeDegradesToSolid: with no fillStyleLst parsed — a
// theme-less synthetic package — idx>1 must fall back to a solid of the
// reference colour rather than dropping the fill.
func TestFillRefWithoutThemeDegradesToSolid(t *testing.T) {
	// A fmtScheme whose fillStyleLst is empty: parseThemeFormatScheme finds
	// nothing, so the style bodies are unknown.
	theme := themeTestTheme("", themeOfficeLnStyles)
	pres := themeTestRead(t, theme, styledShape(
		`<a:fillRef idx="2"><a:schemeClr val="dk1"/></a:fillRef>`, ""), "")
	sh := pres.GetAllSlides()[0].GetShapes()
	if len(sh) == 0 {
		t.Fatal("the slide has no shapes")
	}
	rt, ok := sh[0].(*RichTextShape)
	if !ok {
		t.Fatalf("shape is %T, want *RichTextShape", sh[0])
	}
	if rt.fill == nil || rt.fill.Type != FillSolid || rt.fill.Color.ARGB != "FF000000" {
		t.Fatalf("theme-less fillRef resolved to %+v, want solid FF000000", rt.fill)
	}
}

// TestHyperlinkRunTakesThemeLinkColour: PowerPoint paints a linked run in the
// theme's hlink colour even when the run's rPr declares a solidFill — the
// comparison deck writes tx1 on every link and renders 0000FF. hlinkClick sits
// after solidFill in CT_TextCharacterProperties, so the override lands after
// the fill was parsed.
func TestHyperlinkRunTakesThemeLinkColour(t *testing.T) {
	theme := themeTestTheme(themeOfficeFillStyles, themeOfficeLnStyles)
	shape := `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Linked"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/><a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:hlinkClick r:id="rId1"/></a:rPr><a:t>linked</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`
	rels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://example.com/" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`
	pres := themeTestRead(t, theme, shape, rels)
	sh := pres.GetAllSlides()[0].GetShapes()
	if len(sh) == 0 {
		t.Fatal("the slide has no shapes")
	}
	rt, ok := sh[0].(*RichTextShape)
	if !ok {
		t.Fatalf("shape is %T, want *RichTextShape", sh[0])
	}
	paras := rt.GetParagraphs()
	if len(paras) == 0 {
		t.Fatal("the shape has no paragraphs")
	}
	runs := paras[0].GetElements()
	if len(runs) == 0 {
		t.Fatal("the paragraph has no runs")
	}
	tr, ok := runs[0].(*TextRun)
	if !ok {
		t.Fatalf("element is %T, want *TextRun", runs[0])
	}
	if tr.hyperlink == nil {
		t.Fatal("the run lost its hyperlink")
	}
	if tr.font == nil || tr.font.Color.ARGB != "FF0000FF" {
		got := "nil font"
		if tr.font != nil {
			got = tr.font.Color.ARGB
		}
		t.Errorf("linked run colour = %s, want FF0000FF (the theme hlink; the rPr's tx1 must not win)", got)
	}
	_ = strings.Contains // keep strings imported for future assertions
}
