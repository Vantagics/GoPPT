package gopresentation

import (
	"bytes"
	"encoding/xml"
	"image"
	"strings"
	"testing"
)

// The <a:pPr> tests.
//
// A paragraph's properties were the half of the text pipeline that had drifted
// furthest, in four separate ways:
//
//   - the reader parses marL/marR/indent and the renderer indents the text by
//     all three, but no emitter ever wrote them, so an indented paragraph came
//     back flush against the shape's left edge;
//   - the writer has always emitted lvl and nothing read it, so an outline level
//     could not survive a load;
//   - a table cell built its own <a:p> and put nothing in it but the runs: no
//     <a:pPr> at all, and no <a:br/> either, because only *TextRun was matched;
//   - the second slide reader — the one that walks a layout's non-placeholder
//     shapes, which applyLayoutInheritance then prepends to the slide and draws
//     — understood one attribute of the set, and did not look for <a:br/> at
//     all.
//
// Every assertion below is structural, on the emitted XML text, because a round
// trip cannot see any of it: the reader matches on the local name and does not
// care whether an attribute is there, so a dropped one simply reads back as the
// default, while the renderer draws from the model and the preview therefore
// looked right the whole time.

// pprRun is a plain run, so a fixture's paragraph content stays out of the way.
func pprRun(text string) string {
	return `<a:r><a:rPr lang="en-US" dirty="0" sz="1800"/><a:t>` + text + `</a:t></a:r>`
}

// pprSlideTextBox is a slide carrying one text box whose single paragraph is
// whatever the test is about.
func pprSlideTextBox(pPr, body string) string {
	return `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Para"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>` + pPr + body + `</a:p>
  </p:txBody>
</p:sp>`
}

// pprTableCell is a one-cell table whose paragraph is whatever the test is about.
func pprTableCell(pPr, body string) string {
	return `<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="5" name="Cell"/>
    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm><a:off x="500000" y="500000"/><a:ext cx="6000000" cy="2000000"/></p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblPr firstRow="1" bandRow="1"/>
        <a:tblGrid><a:gridCol w="6000000"/></a:tblGrid>
        <a:tr h="1200000">
          <a:tc>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>` + pPr + body + `</a:p>
            </a:txBody>
            <a:tcPr/>
          </a:tc>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>`
}

// pprLayoutDoc is a slide layout carrying one non-placeholder text shape, which
// is the only kind parseLayoutImages builds a shape from.
func pprLayoutDoc(pPr, body string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="` + nsDrawingML + `" xmlns:r="` + nsOfficeDocRels + `" xmlns:p="` + nsPresentationML + `">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="LayoutText"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>` + pPr + body + `</a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`
}

// layoutShape reads the single text shape pprLayoutDoc describes, the way
// applyLayoutInheritance reads one out of a real layout part.
func layoutShape(t *testing.T, pPr, body string) *RichTextShape {
	t.Helper()
	shapes := (&PPTXReader{}).parseLayoutImages(
		[]byte(pprLayoutDoc(pPr, body)), nil, nil, "ppt/slideLayouts/slideLayout1.xml", New())
	if len(shapes) != 1 {
		t.Fatalf("the layout produced %d shapes, want 1", len(shapes))
	}
	rt, ok := shapes[0].(*RichTextShape)
	if !ok {
		t.Fatalf("the layout shape is %T, want *RichTextShape", shapes[0])
	}
	if len(rt.paragraphs) != 1 {
		t.Fatalf("the layout paragraph count is %d, want 1", len(rt.paragraphs))
	}
	return rt
}

// firstInkColumn returns the x of the left-most inked pixel, or -1 when nothing
// was drawn. Reading a render by eye is not available here, so the ink itself is
// the measurement; the rule is the one countInkIn and inkRowBands use.
func firstInkColumn(img image.Image) int {
	b := img.Bounds()
	for x := b.Min.X; x < b.Max.X; x++ {
		for y := b.Min.Y; y < b.Max.Y; y++ {
			r, g, bl, a := img.At(x, y).RGBA()
			if a>>8 >= 8 && (r>>8 < 250 || g>>8 < 250 || bl>>8 < 250) {
				return x
			}
		}
	}
	return -1
}

// TestApplyPPrAttrsServesAHandBuiltParagraph pins the helper's own contract.
//
// Both scanners hand it a NewParagraph() result, which always carries an
// Alignment, but the helper is the place that knows the attribute block needs
// somewhere to put the values — so it builds one rather than assuming, and
// tolerates being handed nothing at all.
func TestApplyPPrAttrsServesAHandBuiltParagraph(t *testing.T) {
	applyPPrAttrs(nil, []xml.Attr{{Name: xml.Name{Local: "algn"}, Value: "ctr"}})

	p := &Paragraph{}
	applyPPrAttrs(p, []xml.Attr{
		{Name: xml.Name{Local: "algn"}, Value: "ctr"},
		{Name: xml.Name{Local: "marL"}, Value: "123"},
	})
	if p.alignment == nil {
		t.Fatal("the paragraph was left without an alignment to hold the attributes")
	}
	if p.alignment.Horizontal != HorizontalCenter || p.alignment.MarginLeft != 123 {
		t.Errorf("applied algn %q, marL %d; want ctr, 123",
			p.alignment.Horizontal, p.alignment.MarginLeft)
	}
}

// TestParagraphKeepsItsMarginsAndIndent covers marL/marR/indent.
//
// All three are read and all three are used to lay the paragraph out — marL with
// a negative indent is the hanging indent every bulleted list is built from —
// and no emitter wrote a single one of them. The paragraph read back correctly
// and the preview drew it correctly, so only the saved file was wrong.
func TestParagraphKeepsItsMarginsAndIndent(t *testing.T) {
	pres, slide := bodyPrRoundTrip(t, pprSlideTextBox(
		`<a:pPr marL="342900" marR="12700" indent="-342900"/>`, pprRun("INDENTED")))

	rt, ok := firstShape(t, pres).(*RichTextShape)
	if !ok {
		t.Fatalf("shape is %T, want *RichTextShape", firstShape(t, pres))
	}
	if len(rt.paragraphs) != 1 {
		t.Fatalf("the shape has %d paragraphs, want 1", len(rt.paragraphs))
	}
	a := rt.paragraphs[0].alignment
	if a == nil {
		t.Fatal("the paragraph came back with no alignment at all")
	}
	if a.MarginLeft != 342900 || a.MarginRight != 12700 || a.Indent != -342900 {
		t.Errorf("read margins = marL %d, marR %d, indent %d; want 342900, 12700, -342900",
			a.MarginLeft, a.MarginRight, a.Indent)
	}

	got := blockIn(t, slide, "<a:pPr", "</a:pPr>")
	for _, want := range []string{`marL="342900"`, `marR="12700"`, `indent="-342900"`} {
		if !strings.Contains(got, want) {
			t.Errorf("written %s has no %s: the paragraph's indentation was read and "+
				"drawn but never saved", got, want)
		}
	}
}

// TestParagraphLeavesOutMarginsItWasNeverGiven guards the other direction. Zero
// is the schema default for all three, so a paragraph that was never given a
// margin has to be left alone rather than pinned to PowerPoint's default: a
// placeholder inherits its list indentation from the layout, and stating a value
// here would break that inheritance for good.
func TestParagraphLeavesOutMarginsItWasNeverGiven(t *testing.T) {
	_, slide := bodyPrRoundTrip(t, pprSlideTextBox(
		`<a:pPr algn="ctr"/>`, pprRun("CENTERED")))

	got := blockIn(t, slide, "<a:pPr", "</a:pPr>")
	for _, unwanted := range []string{"marL=", "marR=", "indent=", "lvl="} {
		if strings.Contains(got, unwanted) {
			t.Errorf("written %s states %s, but the paragraph was never given one", got, unwanted)
		}
	}
	if !strings.Contains(got, `algn="ctr"`) {
		t.Errorf("written %s lost the alignment the paragraph did carry", got)
	}
}

// TestParagraphLevelIsReadBack covers lvl, the one attribute that went the other
// way: the writer has always emitted it and no reader ever consumed it, so
// Alignment.Level could be set, written, and never read back.
func TestParagraphLevelIsReadBack(t *testing.T) {
	pres, slide := bodyPrRoundTrip(t, pprSlideTextBox(`<a:pPr lvl="2"/>`, pprRun("SUBLEVEL")))

	rt, ok := firstShape(t, pres).(*RichTextShape)
	if !ok {
		t.Fatalf("shape is %T, want *RichTextShape", firstShape(t, pres))
	}
	a := rt.paragraphs[0].alignment
	if a == nil {
		t.Fatal("the paragraph came back with no alignment at all")
	}
	if a.Level != 2 {
		t.Errorf("read level = %d, want 2: the writer emitted lvl and nothing read it back", a.Level)
	}
	if got := blockIn(t, slide, "<a:pPr", "</a:pPr>"); !strings.Contains(got, `lvl="2"`) {
		t.Errorf("written %s has no lvl=\"2\"", got)
	}
}

// TestTableCellKeepsItsParagraphProperties covers the cell's own paragraph
// emitter, which put nothing but the runs into <a:p>.
//
// The reader parses a cell's <a:pPr> — it gates the whole attribute block on
// "inside a text body or inside a cell" — and drawParagraphs draws a cell's
// paragraph through exactly the same path as a shape's, so a cell read from a
// deck was laid out one way in the preview and saved another.
func TestTableCellKeepsItsParagraphProperties(t *testing.T) {
	pres, slide := bodyPrRoundTrip(t, pprTableCell(
		`<a:pPr algn="ctr" marL="228600" indent="-228600">`+
			`<a:spcBef><a:spcPts val="600"/></a:spcBef><a:buChar char="-"/></a:pPr>`,
		pprRun("CELL-A")+`<a:br/>`+pprRun("CELL-B")))

	cell := firstTable(t, pres).GetCell(0, 0)
	if cell == nil {
		t.Fatal("the table has no cell at 0,0")
	}
	if len(cell.paragraphs) != 1 {
		t.Fatalf("the cell has %d paragraphs, want 1", len(cell.paragraphs))
	}
	para := cell.paragraphs[0]

	if a := para.alignment; a == nil {
		t.Error("the reader dropped the cell paragraph's alignment")
	} else if a.Horizontal != HorizontalCenter || a.MarginLeft != 228600 || a.Indent != -228600 {
		t.Errorf("read cell paragraph = algn %q, marL %d, indent %d; want ctr, 228600, -228600",
			a.Horizontal, a.MarginLeft, a.Indent)
	}
	if para.bullet == nil || para.bullet.Type != BulletTypeChar || para.bullet.Style != "-" {
		t.Errorf("read cell bullet = %+v, want a character bullet drawing %q", para.bullet, "-")
	}
	if para.spaceBefore != 600 {
		t.Errorf("read spaceBefore = %d, want 600", para.spaceBefore)
	}
	if n := len(para.elements); n != 3 {
		t.Errorf("the cell paragraph holds %d elements, want 3 (run, break, run)", n)
	}

	tc := blockIn(t, slide, "<a:tc>", "</a:tc>")
	for _, want := range []string{
		"<a:pPr", `algn="ctr"`, `marL="228600"`, `indent="-228600"`,
		`<a:spcBef><a:spcPts val="600"/></a:spcBef>`, `<a:buChar char="-"/>`, "<a:br/>",
	} {
		if !strings.Contains(tc, want) {
			t.Errorf("the written cell has no %s: a cell's paragraph was saved without "+
				"its properties.\n%s", want, tc)
		}
	}
}

// TestTableCellKeepsALineBreakTheModelHolds isolates the writing half of the
// line break: the model holds one, so the emitter has to write one whatever the
// reader did or did not produce.
func TestTableCellKeepsALineBreakTheModelHolds(t *testing.T) {
	pres := packageWithTableXML(t, pprTableCell("", pprRun("FIRST")+pprRun("SECOND")))
	cell := firstTable(t, pres).GetCell(0, 0)
	if cell == nil || len(cell.paragraphs) == 0 {
		t.Fatal("the crafted table has no paragraph to break")
	}
	cell.paragraphs[0].CreateBreak()

	tc := blockIn(t, slidePartOfShapeRound(t, pres), "<a:tc>", "</a:tc>")
	if !strings.Contains(tc, "<a:br/>") {
		t.Errorf("a cell paragraph holding a break was written without one:\n%s", tc)
	}
}

// TestLayoutTextShapeKeepsItsParagraphProperties covers the second slide reader.
//
// parseLayoutImages walks the same markup as parseSlideXML for the
// non-placeholder shapes a layout contributes, and applyLayoutInheritance
// prepends what it returns to the slide — so those shapes are drawn. Its
// paragraph handling understood one attribute of the set.
func TestLayoutTextShapeKeepsItsParagraphProperties(t *testing.T) {
	rt := layoutShape(t,
		`<a:pPr marL="342900" marR="12700" indent="-342900" lvl="1"/>`,
		pprRun("LAYOUT-A")+`<a:br/>`+pprRun("LAYOUT-B"))

	a := rt.paragraphs[0].alignment
	if a == nil {
		t.Fatal("the layout paragraph came back with no alignment at all")
	}
	if a.MarginLeft != 342900 || a.MarginRight != 12700 || a.Indent != -342900 || a.Level != 1 {
		t.Errorf("read layout paragraph = marL %d, marR %d, indent %d, lvl %d; "+
			"want 342900, 12700, -342900, 1", a.MarginLeft, a.MarginRight, a.Indent, a.Level)
	}
	if n := len(rt.paragraphs[0].elements); n != 3 {
		t.Errorf("the layout paragraph holds %d elements, want 3 (run, break, run)", n)
	}
}

// TestLayoutTextShapeIsIndentedWhenDrawn is the rendering half of the same
// defect: the shapes the layout reader builds are drawn, and the paragraph's
// marL is what pushes the text in from the shape's left edge. The measurement is
// where the left-most ink lands, compared between an indented paragraph and an
// unindented one, so the assertion has a direction rather than a pixel count.
func TestLayoutTextShapeIsIndentedWhenDrawn(t *testing.T) {
	fc := NewFontCache()
	requireAnyFont(t, fc)

	leftMostInk := func(marL string) int {
		pPr := `<a:pPr/>`
		if marL != "0" {
			pPr = `<a:pPr marL="` + marL + `"/>`
		}
		rt := layoutShape(t, pPr, pprRun("INDENTED"))

		pres := New()
		pres.GetActiveSlide().AddShape(rt)
		opts := goldenOptions(fc)
		opts.Width = 800
		img, err := pres.SlideToImage(0, opts)
		if err != nil {
			t.Fatalf("render with marL=%s: %v", marL, err)
		}
		x := firstInkColumn(img)
		if x < 0 {
			t.Fatalf("marL=%s drew nothing at all, so the measurement is vacuous", marL)
		}
		return x
	}

	flush := leftMostInk("0")
	indented := leftMostInk("342900")
	t.Logf("left-most ink: %d px without marL, %d px with marL=342900", flush, indented)
	if indented <= flush {
		t.Errorf("the left-most inked column is %d with marL=342900 and %d without: a "+
			"paragraph's indent is not reaching the renderer", indented, flush)
	}
}

// TestParagraphWithoutAlignmentStillSaves covers the public setter that takes
// nil. The renderer has always asked whether a paragraph has an Alignment before
// using it; the writer dereferenced it, so SetAlignment(nil) made WriteTo panic
// out of the library and the deck could not be saved at all.
func TestParagraphWithoutAlignmentStillSaves(t *testing.T) {
	pres, _ := bodyPrRoundTrip(t, pprSlideTextBox(`<a:pPr algn="ctr"/>`, pprRun("TXT")))
	rt, ok := firstShape(t, pres).(*RichTextShape)
	if !ok {
		t.Fatalf("shape is %T, want *RichTextShape", firstShape(t, pres))
	}
	rt.paragraphs[0].SetAlignment(nil)

	var buf bytes.Buffer
	panicked := false
	func() {
		defer func() {
			if r := recover(); r != nil {
				panicked = true
				t.Errorf("WriteTo panicked on a paragraph with no alignment: %v", r)
			}
		}()
		w, err := NewWriter(pres, WriterPowerPoint2007)
		if err != nil {
			t.Fatalf("new writer: %v", err)
		}
		if err := w.WriteTo(&buf); err != nil {
			t.Errorf("writing a paragraph with no alignment failed: %v", err)
		}
	}()

	// A panic leaves a half-written package behind, so there is nothing to
	// inspect unless the save ran to the end.
	if panicked || buf.Len() == 0 {
		return
	}
	slide := string(zipParts(t, buf.Bytes())["ppt/slides/slide1.xml"])
	if got := blockIn(t, slide, "<a:pPr", "</a:pPr>"); strings.Contains(got, "algn=") {
		t.Errorf("written %s states an alignment for a paragraph that has none", got)
	}
}
