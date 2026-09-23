package gopresentation

// A table style's <a:tcBdr> lines are the white grid of every stock
// PowerPoint table. slide31 of 00022823 names tableStyleId
// {5C22544A-...} (Medium Style 2 - Accent 1) and declares no cell border of
// its own — the white separators live entirely in ppt/tableStyles.xml, in the
// wholeTbl part's tcBdr (six 1pt lt1 lines) and firstRow's 3pt bottom. The
// reader used to take the style's fills but drop its lines, so every styled
// table rendered as solid colour blocks butted against each other. The
// tcTxStyle colour (firstRow says lt1 — white header text) was dropped too,
// painting headers black.

import (
	"bytes"
	"testing"
)

const tableBorderStylesFixture = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">
 <a:tblStyle styleId="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" styleName="Medium Style 2 - Accent 1">
  <a:wholeTbl>
   <a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="dk1"/></a:tcTxStyle>
   <a:tcStyle>
    <a:tcBdr>
     <a:left><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:left>
     <a:right><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:right>
     <a:top><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:top>
     <a:bottom><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom>
     <a:insideH><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:insideH>
     <a:insideV><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:insideV>
    </a:tcBdr>
    <a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill>
   </a:tcStyle>
  </a:wholeTbl>
  <a:band1H><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1H>
  <a:band2H><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2H>
  <a:firstRow>
   <a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle>
   <a:tcStyle>
    <a:tcBdr><a:bottom><a:ln w="38100" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom></a:tcBdr>
    <a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill>
   </a:tcStyle>
  </a:firstRow>
 </a:tblStyle>
</a:tblStyleLst>`

// A two-column, two-row table with the header flag on and no cell border of
// its own — the styled table of the fidelity deck, in miniature.
const tableBorderSlideFixture = `
<p:graphicFrame>
  <p:nvGraphicFramePr><p:cNvPr id="2" name="T"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
  <p:xfrm><a:off x="500000" y="500000"/><a:ext cx="4000000" cy="2000000"/></p:xfrm>
  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
   <a:tbl>
    <a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>
    <a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>
    <a:tr h="1000000">
     <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>H1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
     <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>H2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
    </a:tr>
    <a:tr h="1000000">
     <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>D1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
     <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>D2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
    </a:tr>
   </a:tbl>
  </a:graphicData></a:graphic>
</p:graphicFrame>`

func readTableBorderFixture(t *testing.T) *Presentation {
	t.Helper()
	parts := zipParts(t, writeToBytes(t, New()))
	parts["ppt/tableStyles.xml"] = []byte(tableBorderStylesFixture)
	parts["ppt/slides/slide1.xml"] = []byte(bodyPrSlide(tableBorderSlideFixture))
	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the package: %v", err)
	}
	return pres
}

func firstTableShape(t *testing.T, pres *Presentation) *TableShape {
	t.Helper()
	for _, slide := range pres.GetAllSlides() {
		for _, shape := range flattenShapes(slide.GetShapes()) {
			if tbl, ok := shape.(*TableShape); ok {
				return tbl
			}
		}
	}
	t.Fatal("the presentation has no table")
	return nil
}

func assertBorder(t *testing.T, what string, b *Border, style BorderStyle, width int, argb string) {
	t.Helper()
	if b.Style != style {
		t.Errorf("%s style = %s, want %s", what, b.Style, style)
	}
	if style != BorderNone {
		if b.Width != width {
			t.Errorf("%s width = %d, want %d", what, b.Width, width)
		}
		if b.Color.ARGB != argb {
			t.Errorf("%s colour = %s, want %s", what, b.Color.ARGB, argb)
		}
	}
}

// TestParseTableStylesReadsBorders: the tcBdr lines land in the style parts —
// width in points, the scheme reference, and an explicit <a:noFill/> kept
// distinct from an absent element.
func TestParseTableStylesReadsBorders(t *testing.T) {
	styles, _ := parseTableStyles([]byte(tableBorderStylesFixture))
	st := styles["{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"]
	if st == nil {
		t.Fatal("style not parsed")
	}
	insideH := st.wholeTbl.borders.insideH
	if !insideH.declared {
		t.Fatal("wholeTbl insideH not declared")
	}
	if insideH.width != 1 || insideH.scheme != "lt1" {
		t.Errorf("wholeTbl insideH = %+v, want width 1 scheme lt1", insideH)
	}
	firstBottom := st.firstRow.borders.bottom
	if !firstBottom.declared || firstBottom.width != 3 || firstBottom.scheme != "lt1" {
		t.Errorf("firstRow bottom = %+v, want width 3 scheme lt1", firstBottom)
	}
	if st.band1H.borders.insideH.declared {
		t.Error("band1H declares no insideH; the wholeTbl line must show through")
	}
	if st.band1H.borders.top.declared {
		t.Error("band1H's empty tcBdr must not count as a top declaration")
	}
}

// TestApplyTableStyleBordersPaintsWhiteGrid: untouched cells come out with
// the style's lines — white 1pt all round, and the heavier 3pt separator the
// firstRow part pins under the header.
func TestApplyTableStyleBordersPaintsWhiteGrid(t *testing.T) {
	pres := readTableBorderFixture(t)
	tbl := firstTableShape(t, pres)
	white := "FFFFFFFF"

	h := tbl.rows[0][0]
	assertBorder(t, "header top", h.border.Top, BorderSolid, 1, white)
	assertBorder(t, "header left", h.border.Left, BorderSolid, 1, white)
	assertBorder(t, "header bottom (firstRow's 3pt separator)", h.border.Bottom, BorderSolid, 3, white)

	d := tbl.rows[1][0]
	// The edge below the header is shared: the 3pt declaration wins over the
	// body row's 1pt insideH.
	assertBorder(t, "body top (header's heavier separator)", d.border.Top, BorderSolid, 3, white)
	assertBorder(t, "body bottom", d.border.Bottom, BorderSolid, 1, white)
	assertBorder(t, "body left", d.border.Left, BorderSolid, 1, white)
	assertBorder(t, "body right", d.border.Right, BorderSolid, 1, white)
}

// TestApplyTableStyleBordersRespectsExplicitLines: a cell's own tcPr line —
// solid or an explicit noFill — must survive the style.
func TestApplyTableStyleBordersRespectsExplicitLines(t *testing.T) {
	pres := readTableBorderFixture(t)
	tbl := firstTableShape(t, pres)
	d := tbl.rows[1][1]
	red := NewColor("FF0000")
	b := NewBorder()
	b.SetSolidFill(red)
	b.Width = 2
	d.border.Right = b
	d.border.rightDeclared = true
	d.border.Top = NewBorder()
	d.border.Top.SetSolidFill(red)
	d.border.Top.Width = 2
	d.border.topDeclared = true
	d.border.Bottom.Style = BorderNone
	d.border.bottomDeclared = true
	applyTableStyleBorders(tbl, pres)

	if d.border.Right.Color.ARGB != "FFFF0000" {
		t.Errorf("explicit right border colour = %s, want FFFF0000 (untouched by the style)", d.border.Right.Color.ARGB)
	}
	if d.border.Top.Color.ARGB != "FFFF0000" {
		t.Errorf("explicit top border colour = %s, want FFFF0000 (untouched by the style)", d.border.Top.Color.ARGB)
	}
	assertBorder(t, "explicit noFill bottom", d.border.Bottom, BorderNone, 0, "")
	assertBorder(t, "undeclared left still styled", d.border.Left, BorderSolid, 1, "FFFFFFFF")
}

// TestTableStyleTextColorAppliesToHeader: firstRow's tcTxStyle says lt1 —
// the white header text — and runs without a colour of their own take it,
// while a run that declared its own colour keeps it.
func TestTableStyleTextColorAppliesToHeader(t *testing.T) {
	pres := readTableBorderFixture(t)
	tbl := firstTableShape(t, pres)
	for _, ci := range []int{0, 1} {
		for _, elem := range tbl.rows[0][ci].paragraphs[0].elements {
			if tr, ok := elem.(*TextRun); ok {
				if want := (Color{ARGB: "FFFFFFFF"}); tr.font.Color != want {
					t.Errorf("header run colour = %s, want FFFFFFFF (firstRow tcTxStyle lt1)", tr.font.Color.ARGB)
				}
			}
		}
	}
	for _, elem := range tbl.rows[1][0].paragraphs[0].elements {
		if tr, ok := elem.(*TextRun); ok {
			if tr.font.Color.ARGB != "FF000000" {
				t.Errorf("body run colour = %s, want the default black (wholeTbl says dk1)", tr.font.Color.ARGB)
			}
		}
	}
}

// TestTableStyleBordersRoundTrip: the materialised lines survive a
// read-modify-write as explicit tcPr lines.
func TestTableStyleBordersRoundTrip(t *testing.T) {
	pres := readTableBorderFixture(t)
	tbl := firstTableShape(t, pres)
	p := New()
	p.GetActiveSlide().AddShape(tbl)
	data := writeToBytes(t, p)

	pres2, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("read back: %v", err)
	}
	tbl2 := firstTableShape(t, pres2)
	assertBorder(t, "round-tripped header bottom", tbl2.rows[0][0].border.Bottom, BorderSolid, 3, "FFFFFFFF")
}

// TestTableStyleBordersRender: the grid must reach pixels — a white line
// between the two shaded rows where the colour blocks used to butt together.
func TestTableStyleBordersRender(t *testing.T) {
	pres := readTableBorderFixture(t)
	fc := NewFontCache()
	requireAnyFont(t, fc)
	opts := DefaultRenderOptions()
	opts.Width = 720
	opts.FontCache = fc
	img, err := pres.SlideToImage(0, opts)
	if err != nil {
		t.Fatalf("render: %v", err)
	}
	// Table box: x 500000..4500000 EMU, y 500000..2500000 EMU → at 720px
	// across a 10in slide that is x 39..354, y 39..197, rows split at y≈118.
	white := 0
	for y := 115; y <= 121; y++ {
		for x := 50; x < 340; x++ {
			r, g, b, _ := img.At(x, y).RGBA()
			if r>>8 > 240 && g>>8 > 240 && b>>8 > 240 {
				white++
			}
		}
	}
	if white < 200 {
		t.Errorf("white grid pixels at the row separator = %d, want the style's white line painted", white)
	}
}
