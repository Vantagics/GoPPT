package gopresentation

// The r41 table-flag and outline-noFill semantics tests.
//
// Two reader defaults were pinned against the comparison decks:
//
//   - <a:tblPr> flags default to FALSE. NewTableShape turns firstRow/bandRow
//     on for API-created tables, but a file that omits the attribute means
//     false (xsd:boolean has no default). Slide 21 of deck 2 stacks six
//     bandRow-only tables; with the inherited default each grew the style's
//     first-row dark header PowerPoint does not draw.
//   - <a:ln><a:noFill/></a:ln> turns the outline off outright and must keep
//     the <p:style> lnRef fallback suppressed. Deck 1's slide 25 draws theme
//     bands with ln noFill on top of a styled lnRef; the fallback painted a
//     2pt accent line PowerPoint does not draw.

import (
	"testing"
)

const tblFlagsSlide = `
<p:graphicFrame>
  <p:nvGraphicFramePr><p:cNvPr id="2" name="T"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
  <p:xfrm><a:off x="100000" y="100000"/><a:ext cx="2000000" cy="800000"/></p:xfrm>
  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
    <a:tbl>
      <a:tblPr bandRow="1"/>
      <a:tblGrid><a:gridCol w="2000000"/></a:tblGrid>
      <a:tr h="300000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody></a:tc></a:tr>
    </a:tbl>
  </a:graphicData></a:graphic>
</p:graphicFrame>
<p:graphicFrame>
  <p:nvGraphicFramePr><p:cNvPr id="3" name="Bare"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
  <p:xfrm><a:off x="100000" y="1200000"/><a:ext cx="2000000" cy="800000"/></p:xfrm>
  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
    <a:tbl>
      <a:tblGrid><a:gridCol w="2000000"/></a:tblGrid>
      <a:tr h="300000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody></a:tc></a:tr>
    </a:tbl>
  </a:graphicData></a:graphic>
</p:graphicFrame>`

const outlineNoFillSlide = `
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Band"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="200000" y="200000"/><a:ext cx="3000000" cy="300000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="7030A0"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:style>
    <a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>
    <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
    <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
  </p:style>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>`

// TestReaderTableFlagsDefaultFalse: a tblPr that omits a flag must not
// inherit the API-level default — both the bandRow="1" table (firstRow false,
// bandRow true) and the attribute-less table (all false) pin this.
func TestReaderTableFlagsDefaultFalse(t *testing.T) {
	pres := readGradFixture(t, tblFlagsSlide)
	shapes := pres.GetAllSlides()[0].GetShapes()
	var banded, bare *TableShape
	for _, sh := range shapes {
		if ts, ok := sh.(*TableShape); ok {
			if ts.GetName() == "Bare" {
				bare = ts
			} else {
				banded = ts
			}
		}
	}
	if banded == nil || bare == nil {
		t.Fatalf("tables dropped: banded=%v bare=%v", banded != nil, bare != nil)
	}
	if banded.firstRow {
		t.Error("banded firstRow = true, want false (attribute absent)")
	}
	if !banded.bandRow {
		t.Error("banded bandRow = false, want true (bandRow=\"1\")")
	}
	if bare.firstRow || bare.bandRow || bare.firstCol || bare.lastCol {
		t.Errorf("bare table flags = %v/%v/%v/%v, want all false",
			bare.firstRow, bare.bandRow, bare.firstCol, bare.lastCol)
	}
}

// TestAPITableKeepsFirstRowDefault: tables created through the API keep the
// header-row default the writer has always relied on.
func TestAPITableKeepsFirstRowDefault(t *testing.T) {
	tbl := NewTableShape(2, 1)
	if !tbl.firstRow || !tbl.bandRow {
		t.Errorf("API table firstRow=%v bandRow=%v, want true true", tbl.firstRow, tbl.bandRow)
	}
}

// TestOutlineNoFillSuppressesStyleLine: ln noFill must produce a BorderNone
// border, not fall through to the lnRef.
func TestOutlineNoFillSuppressesStyleLine(t *testing.T) {
	pres := readGradFixture(t, outlineNoFillSlide)
	shapes := pres.GetAllSlides()[0].GetShapes()
	if len(shapes) == 0 {
		t.Fatal("shape dropped")
	}
	sh, ok := shapes[0].(interface{ GetBorder() *Border })
	if !ok {
		t.Fatalf("shape %T cannot carry a border", shapes[0])
	}
	b := sh.GetBorder()
	if b == nil || b.Style != BorderNone {
		t.Fatalf("border = %+v, want BorderNone (ln noFill beats the lnRef)", b)
	}
}
