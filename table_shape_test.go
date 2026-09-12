package gopresentation

import (
	"bytes"
	"strings"
	"testing"
)

// The table shape's writer.
//
// A table is the one shape whose content is a grid rather than a rectangle of
// shapes, and its writer had been written by hand for the simple case only. It
// emitted a uniform <a:tblGrid> and a plain <a:tc> per cell, so four things the
// model carries — and the reader parses, and the renderer draws — never reached
// the file:
//
//   - SetColSpan and SetRowSpan, documented in API.md, produced no merge at all;
//   - a cell's borders, which the reader reads from <a:lnL>/<a:lnR>/<a:lnT>/<a:lnB>
//     and renderCellBorders draws, were never written;
//   - the <a:gridCol w> and <a:tr h> values a table read from a file carries were
//     replaced by an even split, resizing every column on a save;
//   - and indexing the grid directly made Save panic on a row shorter than the
//     column count, which is a package the reader accepts without complaint.
//
// The assertions are structural as well as round-tripped: a merge is a property
// of the grid, and a round trip through this library alone cannot see a merge
// that was never written.

// tableOnSlide returns a presentation whose only shape is the given table.
func tableOnSlide(tbl *TableShape) *Presentation {
	p := New()
	tbl.SetOffsetX(500000)
	tbl.SetOffsetY(500000)
	tbl.SetWidth(6000000)
	tbl.SetHeight(2000000)
	p.GetActiveSlide().AddShape(tbl)
	return p
}

// firstTable returns the first table of the presentation.
func firstTable(t *testing.T, pres *Presentation) *TableShape {
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

// tableXML returns the <a:tbl> ... </a:tbl> of the first table in a slide part.
func tableXML(t *testing.T, part []byte) string {
	t.Helper()
	text := string(part)
	open := strings.Index(text, "<a:tbl>")
	end := strings.Index(text, "</a:tbl>")
	if open < 0 || end < open {
		t.Fatalf("no table in the slide part:\n%s", text)
	}
	return text[open : end+len("</a:tbl>")]
}

// tableRowCellCounts returns how many <a:tc> elements each row of a table holds.
//
// The count is what makes a table a grid: <a:tblGrid> declares the columns and
// every row has to fill exactly that many cells, including the empty ones a
// merge leaves behind. A written merge that does not keep the count is a table
// PowerPoint rewrites rather than opens.
func tableRowCellCounts(t *testing.T, table string) []int {
	t.Helper()
	rows := strings.Split(table, "</a:tr>")
	rows = rows[:len(rows)-1] // the remainder after the last row is not a row
	counts := make([]int, 0, len(rows))
	for _, row := range rows {
		// "<a:tc " and "<a:tc>" together, without matching "<a:tcPr>".
		counts = append(counts, strings.Count(row, "<a:tc ")+strings.Count(row, "<a:tc>"))
	}
	return counts
}

func assertCellCounts(t *testing.T, table string, want []int) {
	t.Helper()
	got := tableRowCellCounts(t, table)
	if len(got) != len(want) {
		t.Fatalf("the table has %d rows, want %d: %v", len(got), len(want), got)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Errorf("row %d holds %d cells, want %d: every row must fill the grid\n%s",
				i, got[i], want[i], table)
		}
	}
}

// TestTableCellSpanReachesThePackage pins the merge on its own, without the
// round trip: SetColSpan and SetRowSpan have to change the written grid.
func TestTableCellSpanReachesThePackage(t *testing.T) {
	tbl := NewTableShape(2, 3)
	cell := tbl.GetCell(0, 0)
	cell.SetText("SPANNING")
	cell.SetColSpan(2)
	cell.SetRowSpan(2)

	table := tableXML(t, zipParts(t, writeToBytes(t, tableOnSlide(tbl)))["ppt/slides/slide1.xml"])

	if !strings.Contains(table, `gridSpan="2"`) {
		t.Errorf("SetColSpan(2) wrote no gridSpan:\n%s", table)
	}
	if !strings.Contains(table, `rowSpan="2"`) {
		t.Errorf("SetRowSpan(2) wrote no rowSpan:\n%s", table)
	}
	if !strings.Contains(table, `hMerge="1"`) {
		t.Errorf("the cell to the right of the span is not a continuation:\n%s", table)
	}
	if !strings.Contains(table, `vMerge="1"`) {
		t.Errorf("the cell below the span is not a continuation:\n%s", table)
	}
	if !strings.Contains(table, "SPANNING") {
		t.Errorf("the merged cell's text was lost:\n%s", table)
	}
	// The merge must not be repeated: the spanned cell owns the text and every
	// position it covers other than its own is an empty placeholder.
	if n := strings.Count(table, "SPANNING"); n != 1 {
		t.Errorf("the merged cell's text appears %d times, want 1:\n%s", n, table)
	}

	// The grid keeps its shape: three columns, so three cells in every row.
	assertCellCounts(t, table, []int{3, 3})

	// No column widths were recorded, so the table falls back to an even split.
	if n := strings.Count(table, "<a:gridCol "); n != 3 {
		t.Errorf("the grid declares %d columns, want 3:\n%s", n, table)
	}
	if !strings.Contains(table, `<a:gridCol w="2000000"/>`) {
		t.Errorf("an even split of 6000000 over 3 columns is not 2000000:\n%s", table)
	}
}

// TestTableCellSpanSurvivesRoundTrip is the other half: the merged shape has to
// come back as the merge it was, or a load and a save collapses the table.
func TestTableCellSpanSurvivesRoundTrip(t *testing.T) {
	tbl := NewTableShape(2, 3)
	cell := tbl.GetCell(0, 0)
	cell.SetText("SPANNING")
	cell.SetColSpan(2)
	cell.SetRowSpan(2)

	back := firstTable(t, roundTrip(t, tableOnSlide(tbl)))

	if got := back.GetCell(0, 0).GetColSpan(); got != 2 {
		t.Errorf("column span after round trip = %d, want 2", got)
	}
	if got := back.GetCell(0, 0).GetRowSpan(); got != 2 {
		t.Errorf("row span after round trip = %d, want 2", got)
	}
	if got := back.GetCell(0, 0).GetParagraphs()[0].elements[0].(*TextRun).text; got != "SPANNING" {
		t.Errorf("merged cell text = %q, want %q", got, "SPANNING")
	}
	// The renderer skips a continuation cell, so these two flags decide whether
	// the spanned area is painted twice.
	if c := back.GetCell(0, 1); c == nil || !c.hMerge {
		t.Error("the cell right of the span did not come back as a continuation")
	}
	if c := back.GetCell(1, 0); c == nil || !c.vMerge {
		t.Error("the cell below the span did not come back as a continuation")
	}
}

// TestTableCellBordersReachThePackage covers the other half of <a:tcPr>.
//
// The reader has always read the four <a:lnX> sides and renderCellBorders has
// always drawn them; the writer emitted an empty <a:tcPr>. A deck that set a
// cell's borders therefore lost them on the next save, and a table authored
// through the library could not have one at all.
func TestTableCellBordersReachThePackage(t *testing.T) {
	tbl := NewTableShape(1, 1)
	cell := tbl.GetCell(0, 0)
	cell.SetText("BORDERED")
	borders := cell.GetBorders()
	borders.Top.SetSolidFill(ColorRed).SetWidth(2)
	borders.Left.SetSolidFill(ColorBlue).SetWidth(1)
	cell.SetFill(NewFill().SetSolid(ColorGreen))

	table := tableXML(t, zipParts(t, writeToBytes(t, tableOnSlide(tbl)))["ppt/slides/slide1.xml"])

	// Border.Width is in points; the file wants EMU.
	for _, want := range []string{
		`<a:lnT w="25400">`,
		`<a:lnL w="12700">`,
		`<a:srgbClr val="FF0000"/>`,
		`<a:srgbClr val="0000FF"/>`,
	} {
		if !strings.Contains(table, want) {
			t.Errorf("the cell border has no %s:\n%s", want, table)
		}
	}
	// The two sides nobody set stay out of the file rather than being written
	// as an explicit noFill.
	for _, unwanted := range []string{"<a:lnR", "<a:lnB"} {
		if strings.Contains(table, unwanted) {
			t.Errorf("an unset border side was written as %s:\n%s", unwanted, table)
		}
	}

	// CT_TableCellProperties puts the four line elements before the fill. The
	// comparison is scoped to <a:tcPr>, and the cell's own fill is found by its
	// colour — a border's colour is a <a:solidFill> too, nested inside its
	// <a:lnX>, so the first one in the block belongs to the border.
	pr := table[strings.Index(table, "<a:tcPr>"):]
	lastBorder := strings.LastIndex(pr, "</a:ln")
	fill := strings.Index(pr, `<a:srgbClr val="00FF00"/>`)
	if lastBorder < 0 || fill < 0 || lastBorder > fill {
		t.Errorf("the cell's borders do not come before its fill (last border ends at %d, fill at %d):\n%s", lastBorder, fill, pr)
	}

	// A dash style has to survive as a dash, not decay into a solid line.
	tbl2 := NewTableShape(1, 1)
	tbl2.GetCell(0, 0).GetBorders().Bottom.Style = BorderDash
	dashed := tableXML(t, zipParts(t, writeToBytes(t, tableOnSlide(tbl2)))["ppt/slides/slide1.xml"])
	if !strings.Contains(dashed, `<a:lnB><a:solidFill>`) {
		t.Errorf("the dashed border is malformed:\n%s", dashed)
	}
	if !strings.Contains(dashed, `<a:prstDash val="dash"/>`) {
		t.Errorf("a dashed cell border was written without its dash:\n%s", dashed)
	}
}

// TestTableCellBordersSurviveRoundTrip is the consequence for the reader half:
// the sides have to come back with their width, colour and dash.
func TestTableCellBordersSurviveRoundTrip(t *testing.T) {
	tbl := NewTableShape(1, 1)
	cell := tbl.GetCell(0, 0)
	cell.SetText("BORDERED")
	cell.GetBorders().Top.SetSolidFill(ColorRed).SetWidth(2)
	cell.GetBorders().Bottom.Style = BorderDash
	cell.GetBorders().Bottom.SetWidth(1)

	back := firstTable(t, roundTrip(t, tableOnSlide(tbl))).GetCell(0, 0)
	got := back.GetBorders()

	if got.Top.Style != BorderSolid {
		t.Errorf("top border style = %q, want %q", got.Top.Style, BorderSolid)
	}
	if got.Top.Width != 2 {
		t.Errorf("top border width = %d, want 2 points", got.Top.Width)
	}
	if got.Top.Color.ARGB != ColorRed.ARGB {
		t.Errorf("top border colour = %q, want %q", got.Top.Color.ARGB, ColorRed.ARGB)
	}
	if got.Bottom.Style != BorderDash {
		t.Errorf("bottom border style = %q, want %q: the dash was lost", got.Bottom.Style, BorderDash)
	}
	if got.Right.Style != BorderNone {
		t.Errorf("right border style = %q, want %q: an unset side came back as a line", got.Right.Style, BorderNone)
	}
}

// shortRowTable is the markup of a three-column table whose first row holds two
// cells. The reader takes the column count from <a:tblGrid> and builds rows
// from the <a:tc> elements it finds, so it reads this without complaint.
const shortRowTable = `<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="5" name="Short"/>
    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm><a:off x="500000" y="500000"/><a:ext cx="6000000" cy="2000000"/></p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblPr firstRow="1" bandRow="1"/>
        <a:tblGrid>
          <a:gridCol w="1500000"/>
          <a:gridCol w="4500000"/>
        </a:tblGrid>
        <a:tr h="800000">
          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>SHORT-ROW-A</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
        </a:tr>
        <a:tr h="1200000">
          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>SHORT-ROW-B</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>`

// packageWithTableXML builds a readable package whose slide 1 carries the given
// graphic frame instead of the one the writer produced.
func packageWithTableXML(t *testing.T, frame string) *Presentation {
	t.Helper()
	p := New()
	tbl := NewTableShape(1, 1)
	tbl.SetOffsetX(500000)
	tbl.SetOffsetY(500000)
	tbl.SetWidth(6000000)
	tbl.SetHeight(1200000)
	p.GetActiveSlide().AddShape(tbl)

	parts := zipParts(t, writeToBytes(t, p))
	const part = "ppt/slides/slide1.xml"
	text := string(parts[part])
	open := strings.Index(text, "<p:graphicFrame>")
	end := strings.Index(text, "</p:graphicFrame>")
	if open < 0 || end < open {
		t.Fatalf("the slide has no graphic frame to replace:\n%s", text)
	}
	parts[part] = []byte(text[:open] + frame + text[end+len("</p:graphicFrame>"):])

	pkg := buildZip(t, parts)
	pres, err := (&PPTXReader{}).ReadFromReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil {
		t.Fatalf("read the crafted package: %v", err)
	}
	return pres
}

// TestTableWithAShortRowCanBeSaved is the crash.
//
// The writer indexed the grid as s.rows[i][j] for j < numCols, with numCols
// taken from <a:tblGrid>. A row holding fewer cells than the grid declares is
// read without complaint and drawn without complaint, and then Save panicked —
// recovered into an error by the fail-closed boundary, so the symptom was a
// presentation that could be opened and not written back at all.
func TestTableWithAShortRowCanBeSaved(t *testing.T) {
	pres := packageWithTableXML(t, shortRowTable)

	tbl := firstTable(t, pres)
	if got := tbl.GetNumCols(); got != 2 {
		t.Fatalf("the crafted table has %d columns, want 2", got)
	}
	if got := len(tbl.GetRows()); got != 2 {
		t.Fatalf("the crafted table has %d rows, want 2", got)
	}
	if got := len(tbl.GetRows()[0]); got != 1 {
		t.Fatalf("the crafted first row holds %d cells, want 1 — the fixture is not short", got)
	}

	var buf bytes.Buffer
	if err := pres.WriteTo(&buf); err != nil {
		t.Fatalf("writing a table whose row is shorter than its grid failed: %v", err)
	}

	table := tableXML(t, zipParts(t, buf.Bytes())["ppt/slides/slide1.xml"])
	// One cell per column in every row, the missing one filled in.
	assertCellCounts(t, table, []int{2, 2})
	if !strings.Contains(table, "SHORT-ROW-A") || !strings.Contains(table, "SHORT-ROW-B") {
		t.Errorf("a short row lost its text:\n%s", table)
	}
}

// TestTableGeometrySurvivesRoundTrip covers the last omission in the same
// function: <a:gridCol w> and <a:tr h> were recomputed as an even split, so
// opening and saving a presentation resized every column and row of every table
// in it.
func TestTableGeometrySurvivesRoundTrip(t *testing.T) {
	pres := packageWithTableXML(t, shortRowTable)

	var buf bytes.Buffer
	if err := pres.WriteTo(&buf); err != nil {
		t.Fatalf("write: %v", err)
	}
	table := tableXML(t, zipParts(t, buf.Bytes())["ppt/slides/slide1.xml"])

	for _, want := range []string{
		`<a:gridCol w="1500000"/>`,
		`<a:gridCol w="4500000"/>`,
		`<a:tr h="800000">`,
		`<a:tr h="1200000">`,
	} {
		if !strings.Contains(table, want) {
			t.Errorf("the table geometry is missing %s:\n%s", want, table)
		}
	}
}
