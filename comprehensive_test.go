package gopresentation

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

// helper: write presentation to buffer and read back
func roundTrip(t *testing.T, p *Presentation) *Presentation {
	t.Helper()
	var buf bytes.Buffer
	if err := p.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo failed: %v", err)
	}
	data := buf.Bytes()
	r := bytes.NewReader(data)
	pres, err := ReadFrom(r, int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFrom failed: %v", err)
	}
	return pres
}

// helper: save to temp file and re-open
func roundTripFile(t *testing.T, p *Presentation) *Presentation {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "test.pptx")
	if err := p.Save(path); err != nil {
		t.Fatalf("Save failed: %v", err)
	}
	pres, err := Open(path)
	if err != nil {
		t.Fatalf("Open failed: %v", err)
	}
	return pres
}

// helper: create a minimal 1x1 PNG
func testPNG() []byte {
	return []byte{
		0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
		0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
		0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
		0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
		0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
		0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
		0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
		0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
		0x44, 0xAE, 0x42, 0x60, 0x82,
	}
}


// =============================================================================
// Test 1: Full round-trip with all shape types on a single slide
// =============================================================================
func TestComprehensiveAllShapeTypes(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// 1. RichTextShape with formatting
	rt := slide.CreateRichTextShape()
	rt.SetName("MyTextBox")
	rt.SetPosition(Inch(1), Inch(1))
	rt.SetSize(Inch(4), Inch(1))
	rt.SetWordWrap(true)
	rt.SetColumns(2)
	rt.SetTextAnchor(TextAnchorMiddle)
	rt.SetRotation(45)
	rt.SetFlipHorizontal(true)
	rt.SetDescription("A text box")
	tr := rt.CreateTextRun("Hello World")
	tr.GetFont().SetBold(true).SetItalic(true).SetSize(24).SetColor(ColorRed).SetName("Arial").SetUnderline(UnderlineSingle).SetStrikethrough(true)

	// 2. Second paragraph with different alignment
	para2 := rt.CreateParagraph()
	para2.GetAlignment().SetHorizontal(HorizontalCenter)
	tr2 := para2.CreateTextRun("Centered text")
	tr2.GetFont().SetSize(18).SetColor(ColorBlue)

	// 3. DrawingShape (image)
	img := slide.CreateDrawingShape()
	img.SetImageData(testPNG(), "image/png")
	img.SetName("TestImage")
	img.SetDescription("A test image")
	img.SetPosition(Inch(5), Inch(1))
	img.SetSize(Inch(2), Inch(2))

	// 4. AutoShape
	auto := slide.CreateAutoShape()
	auto.SetAutoShapeType("ellipse")
	auto.SetName("MyEllipse")
	auto.SetPosition(Inch(1), Inch(3))
	auto.SetSize(Inch(2), Inch(1))
	auto.SetText("Inside ellipse")
	auto.GetFill().SetSolid(ColorGreen)

	// 5. LineShape
	line := slide.CreateLineShape()
	line.SetName("MyLine")
	line.SetPosition(Inch(4), Inch(3))
	line.SetSize(Inch(3), 0)
	line.SetLineColor(ColorRed)
	line.SetLineWidth(3)

	// 6. TableShape
	table := slide.CreateTableShape(2, 3)
	table.SetName("MyTable")
	table.SetPosition(Inch(1), Inch(5))
	table.SetSize(Inch(6), Inch(2))
	table.GetCell(0, 0).SetText("Header 1")
	table.GetCell(0, 1).SetText("Header 2")
	table.GetCell(0, 2).SetText("Header 3")
	table.GetCell(1, 0).SetText("Data A")
	table.GetCell(1, 1).SetText("Data B")
	table.GetCell(1, 2).SetText("Data C")
	table.GetCell(0, 0).SetFill(NewFill().SetSolid(ColorBlue))

	// 7. PlaceholderShape
	ph := slide.CreatePlaceholderShape(PlaceholderTitle)
	ph.SetText("Slide Title")
	ph.SetPosition(Inch(0.5), Inch(0))
	ph.SetSize(Inch(8), Inch(1))

	// 8. GroupShape with children
	grp := slide.CreateGroupShape()
	grp.SetName("MyGroup")
	grp.SetPosition(Inch(5), Inch(5))
	grp.SetSize(Inch(3), Inch(2))
	childRT := NewRichTextShape()
	childRT.SetPosition(Inch(5), Inch(5))
	childRT.SetSize(Inch(1), Inch(0.5))
	childRT.CreateTextRun("Group child")
	grp.AddShape(childRT)

	// Validate before saving
	if err := p.Validate(); err != nil {
		t.Fatalf("Validation failed: %v", err)
	}

	// Round-trip
	pres := roundTrip(t, p)

	if pres.GetSlideCount() != 1 {
		t.Fatalf("expected 1 slide, got %d", pres.GetSlideCount())
	}

	s, _ := pres.GetSlide(0)
	shapes := s.GetShapes()

	// Count shape types
	var richTexts, drawings, autos, lines, tables, placeholders, groups int
	for _, sh := range shapes {
		switch sh.(type) {
		case *RichTextShape:
			richTexts++
		case *DrawingShape:
			drawings++
		case *AutoShape:
			autos++
		case *LineShape:
			lines++
		case *TableShape:
			tables++
		case *PlaceholderShape:
			placeholders++
		case *GroupShape:
			groups++
		}
	}

	if richTexts < 1 {
		t.Errorf("expected at least 1 RichTextShape, got %d", richTexts)
	}
	if drawings != 1 {
		t.Errorf("expected 1 DrawingShape, got %d", drawings)
	}
	if autos != 1 {
		t.Errorf("expected 1 AutoShape, got %d", autos)
	}
	if lines != 1 {
		t.Errorf("expected 1 LineShape, got %d", lines)
	}
	if tables != 1 {
		t.Errorf("expected 1 TableShape, got %d", tables)
	}
	if placeholders != 1 {
		t.Errorf("expected 1 PlaceholderShape, got %d", placeholders)
	}
	if groups != 1 {
		t.Errorf("expected 1 GroupShape, got %d", groups)
	}
}


// =============================================================================
// Test 2: Rich text formatting round-trip
// =============================================================================
func TestComprehensiveRichTextFormatting(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetPosition(Inch(1), Inch(1))
	rt.SetSize(Inch(6), Inch(4))
	rt.SetWordWrap(true)
	rt.SetTextAnchor(TextAnchorBottom)

	// Paragraph 1: bold, italic, underline
	tr1 := rt.CreateTextRun("Bold Italic Underline")
	tr1.GetFont().SetBold(true).SetItalic(true).SetUnderline(UnderlineSingle).SetSize(20).SetName("Times New Roman").SetColor(NewColor("FF0000"))

	// Paragraph 2: strikethrough, different font
	para2 := rt.CreateParagraph()
	para2.GetAlignment().SetHorizontal(HorizontalRight)
	para2.SetLineSpacing(200)
	para2.SetSpaceBefore(100)
	para2.SetSpaceAfter(50)
	tr2 := para2.CreateTextRun("Strikethrough text")
	tr2.GetFont().SetStrikethrough(true).SetSize(14).SetColor(ColorBlue)

	// Paragraph 3: break element
	para3 := rt.CreateParagraph()
	para3.CreateTextRun("Before break")
	para3.CreateBreak()
	para3.CreateTextRun("After break")

	pres := roundTrip(t, p)
	s, _ := pres.GetSlide(0)

	// Find the rich text shape
	var found *RichTextShape
	for _, sh := range s.GetShapes() {
		if rts, ok := sh.(*RichTextShape); ok {
			found = rts
			break
		}
	}
	if found == nil {
		t.Fatal("RichTextShape not found after round-trip")
	}

	// Check text anchor
	if found.GetTextAnchor() != TextAnchorBottom {
		t.Errorf("expected text anchor 'b', got '%s'", found.GetTextAnchor())
	}

	// Check word wrap
	if !found.GetWordWrap() {
		t.Error("expected word wrap to be true")
	}

	// Check paragraphs
	paras := found.GetParagraphs()
	if len(paras) < 3 {
		t.Fatalf("expected at least 3 paragraphs, got %d", len(paras))
	}

	// Check first paragraph text run
	elems := paras[0].GetElements()
	if len(elems) == 0 {
		t.Fatal("first paragraph has no elements")
	}
	if tr, ok := elems[0].(*TextRun); ok {
		if tr.GetText() != "Bold Italic Underline" {
			t.Errorf("expected 'Bold Italic Underline', got '%s'", tr.GetText())
		}
		f := tr.GetFont()
		if !f.Bold {
			t.Error("expected bold")
		}
		if !f.Italic {
			t.Error("expected italic")
		}
		if f.Underline != UnderlineSingle {
			t.Errorf("expected underline single, got '%s'", f.Underline)
		}
		if f.Size != 20 {
			t.Errorf("expected font size 20, got %d", f.Size)
		}
		if f.Name != "Times New Roman" {
			t.Errorf("expected font 'Times New Roman', got '%s'", f.Name)
		}
	}

	// Check second paragraph alignment
	if paras[1].GetAlignment().Horizontal != HorizontalRight {
		t.Errorf("expected right alignment, got '%s'", paras[1].GetAlignment().Horizontal)
	}

	// Check strikethrough
	elems2 := paras[1].GetElements()
	if len(elems2) > 0 {
		if tr, ok := elems2[0].(*TextRun); ok {
			if !tr.GetFont().Strikethrough {
				t.Error("expected strikethrough")
			}
		}
	}

	// Check break element in third paragraph
	elems3 := paras[2].GetElements()
	breakCount := 0
	for _, e := range elems3 {
		if e.GetElementType() == "break" {
			breakCount++
		}
	}
	if breakCount != 1 {
		t.Errorf("expected 1 break element, got %d", breakCount)
	}
}


// =============================================================================
// Test 3: Document properties round-trip
// =============================================================================
func TestComprehensiveDocumentProperties(t *testing.T) {
	p := New()
	props := p.GetDocumentProperties()
	props.Creator = "Test Author"
	props.LastModifiedBy = "Test Editor"
	props.Title = "Test Presentation"
	props.Description = "A comprehensive test"
	props.Subject = "Testing"
	props.Keywords = "go, pptx, test"
	props.Category = "Test Category"
	props.Revision = "42"

	// Custom properties
	props.SetCustomProperty("CustomString", "hello", PropertyTypeString)
	props.SetCustomProperty("CustomBool", true, PropertyTypeBoolean)
	props.SetCustomProperty("CustomInt", 42, PropertyTypeInteger)

	pres := roundTrip(t, p)
	rProps := pres.GetDocumentProperties()

	if rProps.Creator != "Test Author" {
		t.Errorf("Creator: expected 'Test Author', got '%s'", rProps.Creator)
	}
	if rProps.Title != "Test Presentation" {
		t.Errorf("Title: expected 'Test Presentation', got '%s'", rProps.Title)
	}
	if rProps.Description != "A comprehensive test" {
		t.Errorf("Description: expected 'A comprehensive test', got '%s'", rProps.Description)
	}
	if rProps.Subject != "Testing" {
		t.Errorf("Subject: expected 'Testing', got '%s'", rProps.Subject)
	}
	if rProps.Keywords != "go, pptx, test" {
		t.Errorf("Keywords: expected 'go, pptx, test', got '%s'", rProps.Keywords)
	}
	if rProps.Category != "Test Category" {
		t.Errorf("Category: expected 'Test Category', got '%s'", rProps.Category)
	}
	if rProps.Revision != "42" {
		t.Errorf("Revision: expected '42', got '%s'", rProps.Revision)
	}
}

// =============================================================================
// Test 4: Multiple slides with different content
// =============================================================================
func TestComprehensiveMultiSlide(t *testing.T) {
	p := New()

	// Slide 1: text
	s1 := p.GetActiveSlide()
	s1.SetName("Slide One")
	s1.SetNotes("Notes for slide 1")
	rt1 := s1.CreateRichTextShape()
	rt1.SetPosition(Inch(1), Inch(1))
	rt1.SetSize(Inch(6), Inch(1))
	rt1.CreateTextRun("Slide 1 content")

	// Slide 2: image
	s2 := p.CreateSlide()
	s2.SetName("Slide Two")
	s2.SetNotes("Notes for slide 2")
	img := s2.CreateDrawingShape()
	img.SetImageData(testPNG(), "image/png")
	img.SetPosition(Inch(2), Inch(2))
	img.SetSize(Inch(3), Inch(3))

	// Slide 3: table
	s3 := p.CreateSlide()
	s3.SetName("Slide Three")
	tbl := s3.CreateTableShape(3, 2)
	tbl.SetPosition(Inch(1), Inch(1))
	tbl.SetSize(Inch(5), Inch(3))
	for r := 0; r < 3; r++ {
		for c := 0; c < 2; c++ {
			tbl.GetCell(r, c).SetText("Cell")
		}
	}

	// Slide 4: chart
	s4 := p.CreateSlide()
	chart := s4.CreateChartShape()
	chart.SetPosition(Inch(1), Inch(1))
	chart.SetSize(Inch(6), Inch(4))
	chart.GetTitle().SetText("Sales Chart")
	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("Q1", []string{"Jan", "Feb", "Mar"}, []float64{10, 20, 30}))
	chart.GetPlotArea().SetType(bar)

	if p.GetSlideCount() != 4 {
		t.Fatalf("expected 4 slides, got %d", p.GetSlideCount())
	}

	pres := roundTripFile(t, p)

	if pres.GetSlideCount() != 4 {
		t.Fatalf("expected 4 slides after round-trip, got %d", pres.GetSlideCount())
	}

	// Verify slide 1 text
	rs1, _ := pres.GetSlide(0)
	text1 := rs1.ExtractText()
	if !strings.Contains(text1, "Slide 1 content") {
		t.Errorf("slide 1 text not found, got: %s", text1)
	}

	// Verify slide 1 notes
	if rs1.GetNotes() != "Notes for slide 1" {
		t.Errorf("slide 1 notes: expected 'Notes for slide 1', got '%s'", rs1.GetNotes())
	}

	// Verify slide 2 has image
	rs2, _ := pres.GetSlide(1)
	hasImage := false
	for _, sh := range rs2.GetShapes() {
		if _, ok := sh.(*DrawingShape); ok {
			hasImage = true
			break
		}
	}
	if !hasImage {
		t.Error("slide 2 should have an image")
	}

	// Verify slide 2 notes
	if rs2.GetNotes() != "Notes for slide 2" {
		t.Errorf("slide 2 notes: expected 'Notes for slide 2', got '%s'", rs2.GetNotes())
	}

	// Verify slide 3 has table
	rs3, _ := pres.GetSlide(2)
	hasTable := false
	for _, sh := range rs3.GetShapes() {
		if ts, ok := sh.(*TableShape); ok {
			hasTable = true
			if ts.GetNumRows() != 3 {
				t.Errorf("table rows: expected 3, got %d", ts.GetNumRows())
			}
			if ts.GetNumCols() != 2 {
				t.Errorf("table cols: expected 2, got %d", ts.GetNumCols())
			}
		}
	}
	if !hasTable {
		t.Error("slide 3 should have a table")
	}
}


// =============================================================================
// Test 5: Comments round-trip
// =============================================================================
func TestComprehensiveComments(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	rt := slide.CreateRichTextShape()
	rt.SetPosition(Inch(1), Inch(1))
	rt.SetSize(Inch(4), Inch(1))
	rt.CreateTextRun("Slide with comments")

	author := NewCommentAuthor("Alice", "A")
	c1 := NewComment()
	c1.SetAuthor(author)
	c1.SetText("First comment")
	c1.SetPosition(100, 200)
	c1.SetDate(time.Date(2025, 6, 15, 10, 30, 0, 0, time.UTC))
	slide.AddComment(c1)

	author2 := NewCommentAuthor("Bob", "B")
	c2 := NewComment()
	c2.SetAuthor(author2)
	c2.SetText("Second comment")
	c2.SetPosition(300, 400)
	slide.AddComment(c2)

	pres := roundTrip(t, p)
	rs, _ := pres.GetSlide(0)

	comments := rs.GetComments()
	if len(comments) != 2 {
		t.Fatalf("expected 2 comments, got %d", len(comments))
	}

	if comments[0].Text != "First comment" {
		t.Errorf("comment 1 text: expected 'First comment', got '%s'", comments[0].Text)
	}
	if comments[0].PositionX != 100 || comments[0].PositionY != 200 {
		t.Errorf("comment 1 position: expected (100,200), got (%d,%d)", comments[0].PositionX, comments[0].PositionY)
	}
	if comments[1].Text != "Second comment" {
		t.Errorf("comment 2 text: expected 'Second comment', got '%s'", comments[1].Text)
	}
}

// =============================================================================
// Test 6: Bullet types round-trip
// =============================================================================
func TestComprehensiveBullets(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetPosition(Inch(1), Inch(1))
	rt.SetSize(Inch(6), Inch(4))

	// Character bullet
	para1 := rt.GetActiveParagraph()
	para1.CreateTextRun("Bullet item 1")
	b1 := NewBullet()
	b1.SetCharBullet("•", "Arial")
	b1.SetColor(ColorRed)
	b1.SetSize(150)
	para1.SetBullet(b1)

	// Numeric bullet
	para2 := rt.CreateParagraph()
	para2.CreateTextRun("Numbered item 1")
	b2 := NewBullet()
	b2.SetNumericBullet(NumFormatArabicPeriod, 5)
	para2.SetBullet(b2)

	// No bullet
	para3 := rt.CreateParagraph()
	para3.CreateTextRun("No bullet")
	b3 := NewBullet()
	b3.Type = BulletTypeNone
	para3.SetBullet(b3)

	pres := roundTrip(t, p)
	rs, _ := pres.GetSlide(0)

	var found *RichTextShape
	for _, sh := range rs.GetShapes() {
		if rts, ok := sh.(*RichTextShape); ok {
			found = rts
			break
		}
	}
	if found == nil {
		t.Fatal("RichTextShape not found")
	}

	paras := found.GetParagraphs()
	if len(paras) < 3 {
		t.Fatalf("expected at least 3 paragraphs, got %d", len(paras))
	}

	// Check character bullet
	if paras[0].GetBullet() == nil {
		t.Fatal("paragraph 1 should have a bullet")
	}
	if paras[0].GetBullet().Type != BulletTypeChar {
		t.Errorf("expected BulletTypeChar, got %d", paras[0].GetBullet().Type)
	}
	if paras[0].GetBullet().Style != "•" {
		t.Errorf("expected bullet char '•', got '%s'", paras[0].GetBullet().Style)
	}
	if paras[0].GetBullet().Font != "Arial" {
		t.Errorf("expected bullet font 'Arial', got '%s'", paras[0].GetBullet().Font)
	}
	if paras[0].GetBullet().Color == nil {
		t.Error("expected bullet color to be set")
	}
	if paras[0].GetBullet().Size != 150 {
		t.Errorf("expected bullet size 150, got %d", paras[0].GetBullet().Size)
	}

	// Check numeric bullet
	if paras[1].GetBullet() == nil {
		t.Fatal("paragraph 2 should have a bullet")
	}
	if paras[1].GetBullet().Type != BulletTypeNumeric {
		t.Errorf("expected BulletTypeNumeric, got %d", paras[1].GetBullet().Type)
	}
	if paras[1].GetBullet().NumFormat != NumFormatArabicPeriod {
		t.Errorf("expected format '%s', got '%s'", NumFormatArabicPeriod, paras[1].GetBullet().NumFormat)
	}
	if paras[1].GetBullet().StartAt != 5 {
		t.Errorf("expected startAt 5, got %d", paras[1].GetBullet().StartAt)
	}

	// Check no bullet
	if paras[2].GetBullet() == nil {
		t.Fatal("paragraph 3 should have a bullet (none type)")
	}
	if paras[2].GetBullet().Type != BulletTypeNone {
		t.Errorf("expected BulletTypeNone, got %d", paras[2].GetBullet().Type)
	}
}


// =============================================================================
// Test 7: All chart types round-trip (write + validate XML)
// =============================================================================
func TestComprehensiveAllChartTypes(t *testing.T) {
	chartTypes := []struct {
		name  string
		chart ChartType
	}{
		{"Bar", NewBarChart()},
		{"Bar3D", NewBar3DChart()},
		{"Line", NewLineChart()},
		{"Pie", NewPieChart()},
		{"Pie3D", NewPie3DChart()},
		{"Doughnut", NewDoughnutChart()},
		{"Scatter", NewScatterChart()},
		{"Radar", NewRadarChart()},
		{"Area", NewAreaChart()},
	}

	series := NewChartSeriesOrdered("Data", []string{"A", "B", "C"}, []float64{10, 20, 30})

	for _, ct := range chartTypes {
		t.Run(ct.name, func(t *testing.T) {
			p := New()
			slide := p.GetActiveSlide()
			chart := slide.CreateChartShape()
			chart.SetPosition(Inch(1), Inch(1))
			chart.SetSize(Inch(6), Inch(4))
			chart.GetTitle().SetText(ct.name + " Chart")

			// Add series to chart type
			switch c := ct.chart.(type) {
			case *BarChart:
				c.AddSeries(series)
			case *Bar3DChart:
				c.BarChart.AddSeries(series)
			case *LineChart:
				c.AddSeries(series)
			case *PieChart:
				c.AddSeries(series)
			case *Pie3DChart:
				c.PieChart.AddSeries(series)
			case *DoughnutChart:
				c.AddSeries(series)
			case *ScatterChart:
				c.AddSeries(series)
			case *RadarChart:
				c.AddSeries(series)
			case *AreaChart:
				c.AddSeries(series)
			}

			chart.GetPlotArea().SetType(ct.chart)

			// Should write without error
			var buf bytes.Buffer
			if err := p.WriteTo(&buf); err != nil {
				t.Fatalf("WriteTo failed for %s chart: %v", ct.name, err)
			}

			// Should be valid zip
			data := buf.Bytes()
			r := bytes.NewReader(data)
			_, err := ReadFrom(r, int64(len(data)))
			if err != nil {
				t.Fatalf("ReadFrom failed for %s chart: %v", ct.name, err)
			}
		})
	}
}

// =============================================================================
// Test 8: Layout types round-trip
// =============================================================================
func TestComprehensiveLayoutTypes(t *testing.T) {
	layouts := []struct {
		name string
		cx   int64
		cy   int64
	}{
		{LayoutScreen4x3, 9144000, 6858000},
		{LayoutScreen16x9, 12192000, 6858000},
		{LayoutScreen16x10, 10972800, 6858000},
		{LayoutA4, 9906000, 6858000},
		{LayoutLetter, 9144000, 6858000},
	}

	for _, l := range layouts {
		t.Run(l.name, func(t *testing.T) {
			p := New()
			p.GetLayout().SetLayout(l.name)

			if p.GetLayout().CX != l.cx {
				t.Errorf("CX: expected %d, got %d", l.cx, p.GetLayout().CX)
			}
			if p.GetLayout().CY != l.cy {
				t.Errorf("CY: expected %d, got %d", l.cy, p.GetLayout().CY)
			}

			pres := roundTrip(t, p)
			if pres.GetLayout().CX != l.cx {
				t.Errorf("after round-trip CX: expected %d, got %d", l.cx, pres.GetLayout().CX)
			}
			if pres.GetLayout().CY != l.cy {
				t.Errorf("after round-trip CY: expected %d, got %d", l.cy, pres.GetLayout().CY)
			}
		})
	}
}

// =============================================================================
// Test 9: Custom layout round-trip
// =============================================================================
func TestComprehensiveCustomLayout(t *testing.T) {
	p := New()
	p.GetLayout().SetCustomLayout(Inch(13.333), Inch(7.5))

	pres := roundTrip(t, p)
	if pres.GetLayout().CX != Inch(13.333) {
		t.Errorf("CX: expected %d, got %d", Inch(13.333), pres.GetLayout().CX)
	}
	if pres.GetLayout().CY != Inch(7.5) {
		t.Errorf("CY: expected %d, got %d", Inch(7.5), pres.GetLayout().CY)
	}
}


// =============================================================================
// Test 10: Slide operations (move, remove, copy)
// =============================================================================
func TestComprehensiveSlideOperations(t *testing.T) {
	p := New()
	s1 := p.GetActiveSlide()
	s1.SetName("First")
	rt1 := s1.CreateRichTextShape()
	rt1.SetPosition(0, 0)
	rt1.SetSize(Inch(1), Inch(1))
	rt1.CreateTextRun("Slide 1")

	s2 := p.CreateSlide()
	s2.SetName("Second")
	rt2 := s2.CreateRichTextShape()
	rt2.SetPosition(0, 0)
	rt2.SetSize(Inch(1), Inch(1))
	rt2.CreateTextRun("Slide 2")

	s3 := p.CreateSlide()
	s3.SetName("Third")
	rt3 := s3.CreateRichTextShape()
	rt3.SetPosition(0, 0)
	rt3.SetSize(Inch(1), Inch(1))
	rt3.CreateTextRun("Slide 3")

	// Test MoveSlide
	if err := p.MoveSlide(2, 0); err != nil {
		t.Fatalf("MoveSlide failed: %v", err)
	}
	first, _ := p.GetSlide(0)
	if first.GetName() != "Third" {
		t.Errorf("after move, first slide should be 'Third', got '%s'", first.GetName())
	}

	// Test RemoveSlideByIndex
	if err := p.RemoveSlideByIndex(1); err != nil {
		t.Fatalf("RemoveSlideByIndex failed: %v", err)
	}
	if p.GetSlideCount() != 2 {
		t.Errorf("expected 2 slides after remove, got %d", p.GetSlideCount())
	}

	// Test CopySlide
	_, err := p.CopySlide(0)
	if err != nil {
		t.Fatalf("CopySlide failed: %v", err)
	}
	if p.GetSlideCount() != 3 {
		t.Errorf("expected 3 slides after copy, got %d", p.GetSlideCount())
	}

	// Verify round-trip
	pres := roundTrip(t, p)
	if pres.GetSlideCount() != 3 {
		t.Errorf("expected 3 slides after round-trip, got %d", pres.GetSlideCount())
	}
}

// =============================================================================
// Test 11: Shape positioning and sizing
// =============================================================================
func TestComprehensiveShapePositioning(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetPosition(Inch(2), Inch(3))
	rt.SetSize(Inch(4), Inch(1.5))
	rt.SetRotation(90)
	rt.SetFlipHorizontal(true)
	rt.SetFlipVertical(true)
	rt.CreateTextRun("Positioned text")

	pres := roundTrip(t, p)
	s, _ := pres.GetSlide(0)

	var found *RichTextShape
	for _, sh := range s.GetShapes() {
		if rts, ok := sh.(*RichTextShape); ok {
			found = rts
			break
		}
	}
	if found == nil {
		t.Fatal("RichTextShape not found")
	}

	if found.GetOffsetX() != Inch(2) {
		t.Errorf("offsetX: expected %d, got %d", Inch(2), found.GetOffsetX())
	}
	if found.GetOffsetY() != Inch(3) {
		t.Errorf("offsetY: expected %d, got %d", Inch(3), found.GetOffsetY())
	}
	if found.GetWidth() != Inch(4) {
		t.Errorf("width: expected %d, got %d", Inch(4), found.GetWidth())
	}
	if found.GetHeight() != Inch(1.5) {
		t.Errorf("height: expected %d, got %d", Inch(1.5), found.GetHeight())
	}
	if found.GetRotation() != 90 {
		t.Errorf("rotation: expected 90, got %d", found.GetRotation())
	}
	if !found.GetFlipHorizontal() {
		t.Error("expected flipH to be true")
	}
	if !found.GetFlipVertical() {
		t.Error("expected flipV to be true")
	}
}

// =============================================================================
// Test 12: Fill and border round-trip
// =============================================================================
func TestComprehensiveFillAndBorder(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// Solid fill shape
	rt := slide.CreateRichTextShape()
	rt.SetPosition(Inch(1), Inch(1))
	rt.SetSize(Inch(3), Inch(1))
	rt.CreateTextRun("Solid fill")
	rt.GetFill().SetSolid(ColorBlue)
	rt.GetBorder().SetSolidFill(ColorRed).SetWidth(25400) // 2pt

	pres := roundTrip(t, p)
	s, _ := pres.GetSlide(0)

	var found *RichTextShape
	for _, sh := range s.GetShapes() {
		if rts, ok := sh.(*RichTextShape); ok {
			found = rts
			break
		}
	}
	if found == nil {
		t.Fatal("RichTextShape not found")
	}

	// Check fill was read back
	if found.GetFill().Type != FillSolid {
		t.Errorf("expected solid fill, got type %d", found.GetFill().Type)
	}
	if found.GetFill().Color.ARGB != ColorBlue.ARGB {
		t.Errorf("fill color: expected %s, got %s", ColorBlue.ARGB, found.GetFill().Color.ARGB)
	}
}


// =============================================================================
// Test 13: Background round-trip
// =============================================================================
func TestComprehensiveBackground(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	rt := slide.CreateRichTextShape()
	rt.SetPosition(Inch(1), Inch(1))
	rt.SetSize(Inch(4), Inch(1))
	rt.CreateTextRun("Slide with background")

	bg := NewFill()
	bg.SetSolid(NewColor("FF8800"))
	slide.SetBackground(bg)

	pres := roundTrip(t, p)
	rs, _ := pres.GetSlide(0)

	if rs.GetBackground() == nil {
		t.Fatal("expected background to be set")
	}
	if rs.GetBackground().Type != FillSolid {
		t.Errorf("expected solid background, got type %d", rs.GetBackground().Type)
	}
}

// =============================================================================
// Test 14: Table with cell fills round-trip
// =============================================================================
func TestComprehensiveTableCellFills(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	tbl := slide.CreateTableShape(2, 2)
	tbl.SetPosition(Inch(1), Inch(1))
	tbl.SetSize(Inch(4), Inch(2))

	tbl.GetCell(0, 0).SetText("Red cell")
	tbl.GetCell(0, 0).SetFill(NewFill().SetSolid(ColorRed))
	tbl.GetCell(0, 1).SetText("Blue cell")
	tbl.GetCell(0, 1).SetFill(NewFill().SetSolid(ColorBlue))
	tbl.GetCell(1, 0).SetText("No fill")
	tbl.GetCell(1, 1).SetText("Green cell")
	tbl.GetCell(1, 1).SetFill(NewFill().SetSolid(ColorGreen))

	pres := roundTrip(t, p)
	rs, _ := pres.GetSlide(0)

	var found *TableShape
	for _, sh := range rs.GetShapes() {
		if ts, ok := sh.(*TableShape); ok {
			found = ts
			break
		}
	}
	if found == nil {
		t.Fatal("TableShape not found")
	}

	if found.GetNumRows() != 2 || found.GetNumCols() != 2 {
		t.Fatalf("table size: expected 2x2, got %dx%d", found.GetNumRows(), found.GetNumCols())
	}

	// Check cell text
	cell00 := found.GetCell(0, 0)
	paras := cell00.GetParagraphs()
	if len(paras) == 0 {
		t.Fatal("cell (0,0) has no paragraphs")
	}
	elems := paras[0].GetElements()
	if len(elems) == 0 {
		t.Fatal("cell (0,0) paragraph has no elements")
	}
	if tr, ok := elems[0].(*TextRun); ok {
		if tr.GetText() != "Red cell" {
			t.Errorf("cell (0,0) text: expected 'Red cell', got '%s'", tr.GetText())
		}
	}
}

// =============================================================================
// Test 15: Image data round-trip
// =============================================================================
func TestComprehensiveImageData(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	imgData := testPNG()
	img := slide.CreateDrawingShape()
	img.SetImageData(imgData, "image/png")
	img.SetName("TestPNG")
	img.SetDescription("A 1x1 PNG")
	img.SetPosition(Inch(2), Inch(2))
	img.SetSize(Inch(3), Inch(3))

	pres := roundTrip(t, p)
	rs, _ := pres.GetSlide(0)

	var found *DrawingShape
	for _, sh := range rs.GetShapes() {
		if ds, ok := sh.(*DrawingShape); ok {
			found = ds
			break
		}
	}
	if found == nil {
		t.Fatal("DrawingShape not found")
	}

	if found.GetImageData() == nil {
		t.Fatal("image data is nil after round-trip")
	}
	if len(found.GetImageData()) != len(imgData) {
		t.Errorf("image data length: expected %d, got %d", len(imgData), len(found.GetImageData()))
	}
	if found.GetMimeType() != "image/png" {
		t.Errorf("mime type: expected 'image/png', got '%s'", found.GetMimeType())
	}
	if found.GetName() != "TestPNG" {
		t.Errorf("name: expected 'TestPNG', got '%s'", found.GetName())
	}
	if found.GetDescription() != "A 1x1 PNG" {
		t.Errorf("description: expected 'A 1x1 PNG', got '%s'", found.GetDescription())
	}
}

// =============================================================================
// Test 16: Line shape round-trip
// =============================================================================
func TestComprehensiveLineShape(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	line := slide.CreateLineShape()
	line.SetName("TestLine")
	line.SetPosition(Inch(1), Inch(1))
	line.SetSize(Inch(5), Inch(0))
	line.SetLineColor(NewColor("00FF00"))
	line.SetLineWidth(4)

	pres := roundTrip(t, p)
	rs, _ := pres.GetSlide(0)

	var found *LineShape
	for _, sh := range rs.GetShapes() {
		if ls, ok := sh.(*LineShape); ok {
			found = ls
			break
		}
	}
	if found == nil {
		t.Fatal("LineShape not found")
	}

	if found.GetName() != "TestLine" {
		t.Errorf("name: expected 'TestLine', got '%s'", found.GetName())
	}
	if found.GetOffsetX() != Inch(1) {
		t.Errorf("offsetX: expected %d, got %d", Inch(1), found.GetOffsetX())
	}
	if found.GetWidth() != Inch(5) {
		t.Errorf("width: expected %d, got %d", Inch(5), found.GetWidth())
	}
	if found.GetLineColor().ARGB != "FF00FF00" {
		t.Errorf("line color: expected 'FF00FF00', got '%s'", found.GetLineColor().ARGB)
	}
	if found.GetLineWidth() != 4 {
		t.Errorf("line width: expected 4, got %d", found.GetLineWidth())
	}
}


// =============================================================================
// Test 17: Placeholder round-trip
// =============================================================================
func TestComprehensivePlaceholder(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	ph := slide.CreatePlaceholderShape(PlaceholderTitle)
	ph.SetText("My Title")
	ph.SetPosition(Inch(0.5), Inch(0.5))
	ph.SetSize(Inch(8), Inch(1))
	ph.SetPlaceholderIndex(0)

	phBody := slide.CreatePlaceholderShape(PlaceholderBody)
	phBody.SetText("Body content")
	phBody.SetPosition(Inch(0.5), Inch(2))
	phBody.SetSize(Inch(8), Inch(4))
	phBody.SetPlaceholderIndex(1)

	pres := roundTrip(t, p)
	rs, _ := pres.GetSlide(0)

	phs := rs.GetPlaceholders()
	if len(phs) != 2 {
		t.Fatalf("expected 2 placeholders, got %d", len(phs))
	}

	titlePH := rs.GetPlaceholder(PlaceholderTitle)
	if titlePH == nil {
		t.Fatal("title placeholder not found")
	}
	if titlePH.GetPlaceholderType() != PlaceholderTitle {
		t.Errorf("expected title type, got '%s'", titlePH.GetPlaceholderType())
	}

	bodyPH := rs.GetPlaceholder(PlaceholderBody)
	if bodyPH == nil {
		t.Fatal("body placeholder not found")
	}

	// Check text content
	text := rs.ExtractText()
	if !strings.Contains(text, "My Title") {
		t.Errorf("expected 'My Title' in text, got: %s", text)
	}
	if !strings.Contains(text, "Body content") {
		t.Errorf("expected 'Body content' in text, got: %s", text)
	}
}

// =============================================================================
// Test 18: Group shape with multiple children round-trip
// =============================================================================
func TestComprehensiveGroupShape(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	grp := slide.CreateGroupShape()
	grp.SetName("TestGroup")
	grp.SetPosition(Inch(1), Inch(1))
	grp.SetSize(Inch(6), Inch(4))

	// Add rich text child
	child1 := NewRichTextShape()
	child1.SetName("GroupText")
	child1.SetPosition(Inch(1), Inch(1))
	child1.SetSize(Inch(2), Inch(1))
	child1.CreateTextRun("Group text child")
	grp.AddShape(child1)

	// Add auto shape child
	child2 := NewAutoShape()
	child2.SetName("GroupEllipse")
	child2.SetAutoShapeType("ellipse")
	child2.SetPosition(Inch(3), Inch(1))
	child2.SetSize(Inch(2), Inch(1))
	grp.AddShape(child2)

	// Add line child
	child3 := NewLineShape()
	child3.SetName("GroupLine")
	child3.SetPosition(Inch(1), Inch(3))
	child3.SetSize(Inch(4), 0)
	child3.SetLineColor(ColorBlue)
	child3.SetLineWidth(2)
	grp.AddShape(child3)

	pres := roundTrip(t, p)
	rs, _ := pres.GetSlide(0)

	var found *GroupShape
	for _, sh := range rs.GetShapes() {
		if gs, ok := sh.(*GroupShape); ok {
			found = gs
			break
		}
	}
	if found == nil {
		t.Fatal("GroupShape not found")
	}

	if found.GetName() != "TestGroup" {
		t.Errorf("group name: expected 'TestGroup', got '%s'", found.GetName())
	}
	if found.GetShapeCount() != 3 {
		t.Errorf("expected 3 children, got %d", found.GetShapeCount())
	}

	// Check child types
	children := found.GetShapes()
	var hasRT, hasAuto, hasLine bool
	for _, ch := range children {
		switch ch.(type) {
		case *RichTextShape:
			hasRT = true
		case *AutoShape:
			hasAuto = true
		case *LineShape:
			hasLine = true
		}
	}
	if !hasRT {
		t.Error("expected RichTextShape child in group")
	}
	if !hasAuto {
		t.Error("expected AutoShape child in group")
	}
	if !hasLine {
		t.Error("expected LineShape child in group")
	}
}

// =============================================================================
// Test 19: Presentation properties round-trip
// =============================================================================
func TestComprehensivePresentationProperties(t *testing.T) {
	p := New()
	pp := p.GetPresentationProperties()
	pp.SetZoom(1.5)
	pp.SetLastView(ViewNotes)
	pp.SetSlideshowType(SlideshowTypeBrowse)

	pres := roundTrip(t, p)
	// Note: presentation properties are not read back from the file
	// (the reader doesn't parse presProps.xml or viewProps.xml)
	// So we just verify the write doesn't fail
	_ = pres
}


// =============================================================================
// Test 20: AutoShape with text and fill round-trip
// =============================================================================
func TestComprehensiveAutoShape(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	auto := slide.CreateAutoShape()
	auto.SetAutoShapeType("ellipse")
	auto.SetName("TestEllipse")
	auto.SetPosition(Inch(2), Inch(2))
	auto.SetSize(Inch(3), Inch(2))
	auto.SetText("Ellipse text")
	auto.GetFill().SetSolid(ColorYellow)

	pres := roundTrip(t, p)
	rs, _ := pres.GetSlide(0)

	var found *AutoShape
	for _, sh := range rs.GetShapes() {
		if as, ok := sh.(*AutoShape); ok {
			found = as
			break
		}
	}
	if found == nil {
		t.Fatal("AutoShape not found")
	}

	if found.GetAutoShapeType() != "ellipse" {
		t.Errorf("shape type: expected 'ellipse', got '%s'", found.GetAutoShapeType())
	}
	if found.GetText() != "Ellipse text" {
		t.Errorf("text: expected 'Ellipse text', got '%s'", found.GetText())
	}
}

// =============================================================================
// Test 21: ExtractText across all shape types
// =============================================================================
func TestComprehensiveExtractText(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// RichText
	rt := slide.CreateRichTextShape()
	rt.SetPosition(0, 0)
	rt.SetSize(Inch(1), Inch(1))
	rt.CreateTextRun("RichText content")

	// Placeholder
	ph := slide.CreatePlaceholderShape(PlaceholderTitle)
	ph.SetText("Placeholder content")
	ph.SetPosition(0, Inch(1))
	ph.SetSize(Inch(1), Inch(1))

	// AutoShape
	auto := slide.CreateAutoShape()
	auto.SetAutoShapeType("ellipse")
	auto.SetPosition(0, Inch(2))
	auto.SetSize(Inch(1), Inch(1))
	auto.SetText("AutoShape content")

	// Table
	tbl := slide.CreateTableShape(1, 1)
	tbl.SetPosition(0, Inch(3))
	tbl.SetSize(Inch(1), Inch(1))
	tbl.GetCell(0, 0).SetText("Table content")

	text := slide.ExtractText()
	for _, expected := range []string{"RichText content", "Placeholder content", "AutoShape content", "Table content"} {
		if !strings.Contains(text, expected) {
			t.Errorf("expected '%s' in extracted text, got: %s", expected, text)
		}
	}

	// Also test presentation-level ExtractText
	pText := p.ExtractText()
	if !strings.Contains(pText, "RichText content") {
		t.Errorf("presentation ExtractText should contain 'RichText content'")
	}
}

// =============================================================================
// Test 22: Validation catches errors
// =============================================================================
func TestComprehensiveValidation(t *testing.T) {
	// Valid presentation
	p := New()
	if err := p.Validate(); err != nil {
		t.Errorf("valid presentation should pass validation: %v", err)
	}

	// Invalid: negative width
	p2 := New()
	slide := p2.GetActiveSlide()
	rt := slide.CreateRichTextShape()
	rt.SetPosition(0, 0)
	rt.SetSize(-100, Inch(1))
	rt.CreateTextRun("test")
	if err := p2.Validate(); err == nil {
		t.Error("expected validation error for negative width")
	}

	// Invalid: drawing with no data
	p3 := New()
	s3 := p3.GetActiveSlide()
	d := s3.CreateDrawingShape()
	d.SetPosition(0, 0)
	d.SetSize(Inch(1), Inch(1))
	if err := p3.Validate(); err == nil {
		t.Error("expected validation error for drawing with no data")
	}

	// Invalid: chart with no type
	p4 := New()
	s4 := p4.GetActiveSlide()
	ch := s4.CreateChartShape()
	ch.SetPosition(0, 0)
	ch.SetSize(Inch(1), Inch(1))
	if err := p4.Validate(); err == nil {
		t.Error("expected validation error for chart with no type")
	}
}

// =============================================================================
// Test 23: Color operations
// =============================================================================
func TestComprehensiveColors(t *testing.T) {
	// 6-char RGB
	c1 := NewColor("FF0000")
	if c1.ARGB != "FFFF0000" {
		t.Errorf("expected 'FFFF0000', got '%s'", c1.ARGB)
	}
	if c1.GetRed() != 255 {
		t.Errorf("red: expected 255, got %d", c1.GetRed())
	}
	if c1.GetGreen() != 0 {
		t.Errorf("green: expected 0, got %d", c1.GetGreen())
	}
	if c1.GetBlue() != 0 {
		t.Errorf("blue: expected 0, got %d", c1.GetBlue())
	}
	if c1.GetAlpha() != 255 {
		t.Errorf("alpha: expected 255, got %d", c1.GetAlpha())
	}

	// 8-char ARGB
	c2 := NewColor("80FF00FF")
	if c2.ARGB != "80FF00FF" {
		t.Errorf("expected '80FF00FF', got '%s'", c2.ARGB)
	}
	if c2.GetAlpha() != 128 {
		t.Errorf("alpha: expected 128, got %d", c2.GetAlpha())
	}

	// With hash prefix
	c3 := NewColor("#00FF00")
	if c3.ARGB != "FF00FF00" {
		t.Errorf("expected 'FF00FF00', got '%s'", c3.ARGB)
	}

	// Invalid color falls back to black
	c4 := NewColor("ZZZZZZ")
	if c4.ARGB != "FF000000" {
		t.Errorf("expected fallback to black, got '%s'", c4.ARGB)
	}

	// Lowercase
	c5 := NewColor("ff0000")
	if c5.ARGB != "FFFF0000" {
		t.Errorf("expected 'FFFF0000', got '%s'", c5.ARGB)
	}
}

// =============================================================================
// Test 24: Measurement conversions
// =============================================================================
func TestComprehensiveMeasurements(t *testing.T) {
	// Inch
	if Inch(1) != 914400 {
		t.Errorf("1 inch = %d EMU, expected 914400", Inch(1))
	}

	// Point
	if Point(1) != 12700 {
		t.Errorf("1 point = %d EMU, expected 12700", Point(1))
	}

	// Centimeter
	if Centimeter(1) != 360000 {
		t.Errorf("1 cm = %d EMU, expected 360000", Centimeter(1))
	}

	// Millimeter
	if Millimeter(1) != 36000 {
		t.Errorf("1 mm = %d EMU, expected 36000", Millimeter(1))
	}

	// Round-trip conversions
	if EMUToInch(Inch(2.5)) != 2.5 {
		t.Errorf("EMUToInch round-trip failed")
	}
	if EMUToPoint(Point(72)) != 72 {
		t.Errorf("EMUToPoint round-trip failed")
	}
	if EMUToCentimeter(Centimeter(10)) != 10 {
		t.Errorf("EMUToCentimeter round-trip failed")
	}
	if EMUToMillimeter(Millimeter(100)) != 100 {
		t.Errorf("EMUToMillimeter round-trip failed")
	}
}


// =============================================================================
// Test 25: Save to file and re-open
// =============================================================================
func TestComprehensiveSaveAndOpen(t *testing.T) {
	p := New()
	props := p.GetDocumentProperties()
	props.Title = "File Test"
	props.Creator = "GoPresentation"

	slide := p.GetActiveSlide()
	rt := slide.CreateRichTextShape()
	rt.SetPosition(Inch(1), Inch(1))
	rt.SetSize(Inch(6), Inch(1))
	tr := rt.CreateTextRun("Saved to file")
	tr.GetFont().SetBold(true).SetSize(24)

	slide.SetNotes("File test notes")

	dir := t.TempDir()
	path := filepath.Join(dir, "test_save.pptx")

	if err := p.Save(path); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Verify file exists
	info, err := os.Stat(path)
	if err != nil {
		t.Fatalf("file not found: %v", err)
	}
	if info.Size() == 0 {
		t.Fatal("file is empty")
	}

	// Re-open
	pres, err := Open(path)
	if err != nil {
		t.Fatalf("Open failed: %v", err)
	}

	if pres.GetDocumentProperties().Title != "File Test" {
		t.Errorf("title: expected 'File Test', got '%s'", pres.GetDocumentProperties().Title)
	}
	if pres.GetSlideCount() != 1 {
		t.Fatalf("expected 1 slide, got %d", pres.GetSlideCount())
	}

	rs, _ := pres.GetSlide(0)
	text := rs.ExtractText()
	if !strings.Contains(text, "Saved to file") {
		t.Errorf("expected 'Saved to file' in text, got: %s", text)
	}
	if rs.GetNotes() != "File test notes" {
		t.Errorf("notes: expected 'File test notes', got '%s'", rs.GetNotes())
	}
}

// =============================================================================
// Test 26: OpenTemplate removes slides
// =============================================================================
func TestComprehensiveOpenTemplate(t *testing.T) {
	// First create a file with 3 slides
	p := New()
	p.CreateSlide()
	p.CreateSlide()

	dir := t.TempDir()
	path := filepath.Join(dir, "template.pptx")
	if err := p.Save(path); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Open as template
	tmpl, err := OpenTemplate(path)
	if err != nil {
		t.Fatalf("OpenTemplate failed: %v", err)
	}

	if tmpl.GetSlideCount() != 0 {
		t.Errorf("template should have 0 slides, got %d", tmpl.GetSlideCount())
	}
}

// =============================================================================
// Test 27: Multiple images on same slide
// =============================================================================
func TestComprehensiveMultipleImages(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	for i := 0; i < 3; i++ {
		img := slide.CreateDrawingShape()
		img.SetImageData(testPNG(), "image/png")
		img.SetPosition(Inch(float64(i)*3), Inch(1))
		img.SetSize(Inch(2), Inch(2))
	}

	pres := roundTrip(t, p)
	rs, _ := pres.GetSlide(0)

	imgCount := 0
	for _, sh := range rs.GetShapes() {
		if _, ok := sh.(*DrawingShape); ok {
			imgCount++
		}
	}
	if imgCount != 3 {
		t.Errorf("expected 3 images, got %d", imgCount)
	}
}

// =============================================================================
// Test 28: Complex presentation with all features combined
// =============================================================================
func TestComprehensiveComplexPresentation(t *testing.T) {
	p := New()
	p.GetDocumentProperties().Title = "Complex Test"
	p.GetDocumentProperties().Creator = "Test Suite"
	p.GetLayout().SetLayout(LayoutScreen16x9)

	// Slide 1: Title slide
	s1 := p.GetActiveSlide()
	s1.SetName("Title Slide")
	s1.SetNotes("This is the title slide")
	s1.SetBackground(NewFill().SetSolid(NewColor("EEEEEE")))

	title := s1.CreatePlaceholderShape(PlaceholderCtrTitle)
	title.SetText("Complex Presentation")
	title.SetPosition(Inch(1), Inch(2))
	title.SetSize(Inch(8), Inch(2))

	subtitle := s1.CreatePlaceholderShape(PlaceholderSubTitle)
	subtitle.SetText("Created by GoPresentation")
	subtitle.SetPosition(Inch(1), Inch(4))
	subtitle.SetSize(Inch(8), Inch(1))

	// Slide 2: Content with bullets
	s2 := p.CreateSlide()
	s2.SetName("Content Slide")
	rt := s2.CreateRichTextShape()
	rt.SetPosition(Inch(1), Inch(1))
	rt.SetSize(Inch(8), Inch(5))
	rt.SetWordWrap(true)

	for i, item := range []string{"First point", "Second point", "Third point"} {
		var para *Paragraph
		if i == 0 {
			para = rt.GetActiveParagraph()
		} else {
			para = rt.CreateParagraph()
		}
		para.CreateTextRun(item)
		b := NewBullet()
		b.SetCharBullet("•")
		para.SetBullet(b)
	}

	// Slide 3: Image + text
	s3 := p.CreateSlide()
	s3.SetName("Image Slide")
	img := s3.CreateDrawingShape()
	img.SetImageData(testPNG(), "image/png")
	img.SetPosition(Inch(1), Inch(1))
	img.SetSize(Inch(4), Inch(4))

	caption := s3.CreateRichTextShape()
	caption.SetPosition(Inch(6), Inch(2))
	caption.SetSize(Inch(3), Inch(1))
	tr := caption.CreateTextRun("Image caption")
	tr.GetFont().SetItalic(true).SetSize(14)

	// Slide 4: Table + chart
	s4 := p.CreateSlide()
	s4.SetName("Data Slide")
	tbl := s4.CreateTableShape(3, 3)
	tbl.SetPosition(Inch(0.5), Inch(0.5))
	tbl.SetSize(Inch(4), Inch(2.5))
	for r := 0; r < 3; r++ {
		for c := 0; c < 3; c++ {
			tbl.GetCell(r, c).SetText("Data")
		}
	}

	chart := s4.CreateChartShape()
	chart.SetPosition(Inch(5), Inch(0.5))
	chart.SetSize(Inch(4.5), Inch(3))
	chart.GetTitle().SetText("Revenue")
	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("2024", []string{"Q1", "Q2", "Q3", "Q4"}, []float64{100, 150, 200, 180}))
	bar.AddSeries(NewChartSeriesOrdered("2025", []string{"Q1", "Q2", "Q3", "Q4"}, []float64{120, 170, 220, 210}))
	chart.GetPlotArea().SetType(bar)

	// Slide 5: Comments
	s5 := p.CreateSlide()
	s5.SetName("Review Slide")
	reviewText := s5.CreateRichTextShape()
	reviewText.SetPosition(Inch(1), Inch(1))
	reviewText.SetSize(Inch(8), Inch(5))
	reviewText.CreateTextRun("Content for review")

	author := NewCommentAuthor("Reviewer", "R")
	comment := NewComment()
	comment.SetAuthor(author).SetText("Please review this").SetPosition(100, 100)
	s5.AddComment(comment)

	// Validate
	if err := p.Validate(); err != nil {
		t.Fatalf("Validation failed: %v", err)
	}

	// Save and re-open
	pres := roundTripFile(t, p)

	if pres.GetSlideCount() != 5 {
		t.Fatalf("expected 5 slides, got %d", pres.GetSlideCount())
	}

	// Verify title
	if pres.GetDocumentProperties().Title != "Complex Test" {
		t.Errorf("title: expected 'Complex Test', got '%s'", pres.GetDocumentProperties().Title)
	}

	// Verify layout
	if pres.GetLayout().CX != 12192000 {
		t.Errorf("layout CX: expected 12192000, got %d", pres.GetLayout().CX)
	}

	// Verify slide 1 background
	rs1, _ := pres.GetSlide(0)
	if rs1.GetBackground() == nil {
		t.Error("slide 1 should have a background")
	}

	// Verify slide 1 notes
	if rs1.GetNotes() != "This is the title slide" {
		t.Errorf("slide 1 notes: expected 'This is the title slide', got '%s'", rs1.GetNotes())
	}

	// Verify slide 5 comments
	rs5, _ := pres.GetSlide(4)
	if rs5.GetCommentCount() != 1 {
		t.Errorf("slide 5 comments: expected 1, got %d", rs5.GetCommentCount())
	}

	// Verify full text extraction
	fullText := pres.ExtractText()
	for _, expected := range []string{"Complex Presentation", "Content for review"} {
		if !strings.Contains(fullText, expected) {
			t.Errorf("expected '%s' in full text", expected)
		}
	}
}


// =============================================================================
// Test 29: Hyperlink round-trip
// =============================================================================
func TestComprehensiveHyperlink(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetPosition(Inch(1), Inch(1))
	rt.SetSize(Inch(6), Inch(1))
	tr := rt.CreateTextRun("Click here")
	tr.SetHyperlink(NewHyperlink("https://example.com"))

	// Verify invalid scheme is rejected
	badLink := NewHyperlink("javascript:alert(1)")
	if badLink != nil {
		t.Error("javascript: scheme should be rejected")
	}

	// Internal hyperlink
	rt2 := slide.CreateRichTextShape()
	rt2.SetPosition(Inch(1), Inch(2))
	rt2.SetSize(Inch(6), Inch(1))
	tr2 := rt2.CreateTextRun("Go to slide 2")
	tr2.SetHyperlink(NewInternalHyperlink(2))

	// Should write without error
	var buf bytes.Buffer
	if err := p.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo failed: %v", err)
	}
}

// =============================================================================
// Test 30: Font clamping
// =============================================================================
func TestComprehensiveFontClamping(t *testing.T) {
	f := NewFont()

	// Size clamping
	f.SetSize(0)
	if f.Size != 1 {
		t.Errorf("expected min size 1, got %d", f.Size)
	}
	f.SetSize(5000)
	if f.Size != 4000 {
		t.Errorf("expected max size 4000, got %d", f.Size)
	}
}

// =============================================================================
// Test 31: Rotation normalization
// =============================================================================
func TestComprehensiveRotation(t *testing.T) {
	b := &BaseShape{}

	b.SetRotation(45)
	if b.GetRotation() != 45 {
		t.Errorf("expected 45, got %d", b.GetRotation())
	}

	b.SetRotation(360)
	if b.GetRotation() != 0 {
		t.Errorf("expected 0 for 360, got %d", b.GetRotation())
	}

	b.SetRotation(-90)
	if b.GetRotation() != 270 {
		t.Errorf("expected 270 for -90, got %d", b.GetRotation())
	}

	b.SetRotation(720)
	if b.GetRotation() != 0 {
		t.Errorf("expected 0 for 720, got %d", b.GetRotation())
	}
}

// =============================================================================
// Test 32: Gradient fill round-trip
// =============================================================================
func TestComprehensiveGradientFill(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetPosition(Inch(1), Inch(1))
	rt.SetSize(Inch(4), Inch(2))
	rt.CreateTextRun("Gradient")
	rt.GetFill().SetGradientLinear(ColorRed, ColorBlue, 90)

	// Verify gradient rotation normalization
	f := NewFill()
	f.SetGradientLinear(ColorRed, ColorBlue, -90)
	if f.Rotation != 270 {
		t.Errorf("expected rotation 270 for -90, got %d", f.Rotation)
	}

	// Should write without error
	var buf bytes.Buffer
	if err := p.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo failed: %v", err)
	}
}

// =============================================================================
// Test 33: Shadow on drawing shape
// =============================================================================
func TestComprehensiveShadow(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	img := slide.CreateDrawingShape()
	img.SetImageData(testPNG(), "image/png")
	img.SetPosition(Inch(2), Inch(2))
	img.SetSize(Inch(3), Inch(3))

	shadow := NewShadow()
	shadow.SetVisible(true).SetDirection(45).SetDistance(5)
	shadow.BlurRadius = 3
	shadow.Alpha = 60
	img.SetShadow(shadow)

	var buf bytes.Buffer
	if err := p.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo failed: %v", err)
	}
}

// =============================================================================
// Test 34: Slide visibility
// =============================================================================
func TestComprehensiveSlideVisibility(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	rt := slide.CreateRichTextShape()
	rt.SetPosition(0, 0)
	rt.SetSize(Inch(1), Inch(1))
	rt.CreateTextRun("test")

	if !slide.IsVisible() {
		t.Error("slide should be visible by default")
	}

	slide.SetVisible(false)
	if slide.IsVisible() {
		t.Error("slide should be hidden")
	}

	slide.SetVisible(true)
	if !slide.IsVisible() {
		t.Error("slide should be visible again")
	}
}

// =============================================================================
// Test 35: Slide transition
// =============================================================================
func TestComprehensiveSlideTransition(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	rt := slide.CreateRichTextShape()
	rt.SetPosition(0, 0)
	rt.SetSize(Inch(1), Inch(1))
	rt.CreateTextRun("test")

	tr := &Transition{
		Type:     TransitionFade,
		Speed:    TransitionSpeedMedium,
		Duration: 500,
	}
	slide.SetTransition(tr)

	got := slide.GetTransition()
	if got == nil {
		t.Fatal("transition should be set")
	}
	if got.Type != TransitionFade {
		t.Errorf("expected TransitionFade, got %d", got.Type)
	}
	if got.Speed != TransitionSpeedMedium {
		t.Errorf("expected medium speed, got '%s'", got.Speed)
	}
	if got.Duration != 500 {
		t.Errorf("expected duration 500, got %d", got.Duration)
	}
}


// =============================================================================
// Test 36: Chart with axes configuration
// =============================================================================
func TestComprehensiveChartAxes(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.SetPosition(Inch(1), Inch(1))
	chart.SetSize(Inch(6), Inch(4))
	chart.GetTitle().SetText("Axis Test")

	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("Data", []string{"A", "B", "C"}, []float64{10, 20, 30}))
	bar.SetBarGrouping("clustered")
	bar.SetGapWidthPercent(150)
	bar.SetOverlapPercent(50)
	chart.GetPlotArea().SetType(bar)

	// Configure axes
	chart.GetPlotArea().GetAxisX().SetTitle("Categories").SetVisible(true)
	chart.GetPlotArea().GetAxisY().SetTitle("Values").SetVisible(true).SetMinBounds(0).SetMaxBounds(50).SetMajorUnit(10)
	chart.GetPlotArea().GetAxisY().SetMajorGridlines(NewGridlines())

	// Legend
	chart.GetLegend().Visible = true
	chart.GetLegend().Position = "b"

	var buf bytes.Buffer
	if err := p.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo failed: %v", err)
	}

	// Read back
	data := buf.Bytes()
	r := bytes.NewReader(data)
	_, err := ReadFrom(r, int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFrom failed: %v", err)
	}
}

// =============================================================================
// Test 37: Chart series with data labels and fill colors
// =============================================================================
func TestComprehensiveChartSeriesOptions(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.SetPosition(Inch(1), Inch(1))
	chart.SetSize(Inch(6), Inch(4))
	chart.GetTitle().SetText("Series Options")
	chart.SetDisplayBlankAs(ChartBlankAsGap)

	series := NewChartSeriesOrdered("Sales", []string{"Jan", "Feb", "Mar"}, []float64{100, 200, 300})
	series.SetFillColor(ColorBlue)
	series.SetLabelPosition("outEnd")

	bar := NewBarChart()
	bar.AddSeries(series)
	chart.GetPlotArea().SetType(bar)

	var buf bytes.Buffer
	if err := p.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo failed: %v", err)
	}
}

// =============================================================================
// Test 38: Paragraph alignment and spacing
// =============================================================================
func TestComprehensiveParagraphAlignment(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	rt := slide.CreateRichTextShape()
	rt.SetPosition(Inch(1), Inch(1))
	rt.SetSize(Inch(6), Inch(4))

	alignments := []HorizontalAlignment{
		HorizontalLeft, HorizontalCenter, HorizontalRight, HorizontalJustify,
	}

	for i, align := range alignments {
		var para *Paragraph
		if i == 0 {
			para = rt.GetActiveParagraph()
		} else {
			para = rt.CreateParagraph()
		}
		para.GetAlignment().SetHorizontal(align)
		para.SetLineSpacing(200)
		para.CreateTextRun("Aligned text")
	}

	pres := roundTrip(t, p)
	rs, _ := pres.GetSlide(0)

	var found *RichTextShape
	for _, sh := range rs.GetShapes() {
		if rts, ok := sh.(*RichTextShape); ok {
			found = rts
			break
		}
	}
	if found == nil {
		t.Fatal("RichTextShape not found")
	}

	paras := found.GetParagraphs()
	if len(paras) < 4 {
		t.Fatalf("expected at least 4 paragraphs, got %d", len(paras))
	}

	for i, align := range alignments {
		if paras[i].GetAlignment().Horizontal != align {
			t.Errorf("paragraph %d: expected alignment '%s', got '%s'", i, align, paras[i].GetAlignment().Horizontal)
		}
	}
}

// =============================================================================
// Test 39: Error handling - invalid operations
// =============================================================================
func TestComprehensiveErrorHandling(t *testing.T) {
	p := New()

	// Out of range slide access
	_, err := p.GetSlide(99)
	if err == nil {
		t.Error("expected error for out-of-range slide")
	}

	// Set active slide out of range
	if err := p.SetActiveSlideIndex(99); err == nil {
		t.Error("expected error for out-of-range active slide")
	}

	// Remove last slide
	if err := p.RemoveSlideByIndex(0); err == nil {
		t.Error("expected error when removing last slide")
	}

	// Move slide out of range
	if err := p.MoveSlide(0, 99); err == nil {
		t.Error("expected error for out-of-range move")
	}

	// Copy slide out of range
	_, err = p.CopySlide(99)
	if err == nil {
		t.Error("expected error for out-of-range copy")
	}

	// Remove shape out of range
	slide := p.GetActiveSlide()
	if err := slide.RemoveShape(99); err == nil {
		t.Error("expected error for out-of-range shape removal")
	}

	// Table cell out of range
	tbl := NewTableShape(2, 2)
	cell := tbl.GetCell(99, 99)
	if cell != nil {
		t.Error("expected nil for out-of-range cell")
	}

	// Read non-existent file
	_, err = Open("nonexistent.pptx")
	if err == nil {
		t.Error("expected error for non-existent file")
	}

	// Read invalid data
	_, err = ReadFrom(bytes.NewReader([]byte("not a zip")), 10)
	if err == nil {
		t.Error("expected error for invalid data")
	}

	// Unsupported reader format
	_, err = NewReader("unsupported")
	if err == nil {
		t.Error("expected error for unsupported reader format")
	}

	// Unsupported writer format
	_, err = NewWriter(p, "unsupported")
	if err == nil {
		t.Error("expected error for unsupported writer format")
	}
}

// =============================================================================
// Test 40: Close and resource cleanup
// =============================================================================
func TestComprehensiveClose(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()
	rt := slide.CreateRichTextShape()
	rt.SetPosition(0, 0)
	rt.SetSize(Inch(1), Inch(1))
	rt.CreateTextRun("test")

	if err := p.Close(); err != nil {
		t.Fatalf("Close failed: %v", err)
	}

	// After close, slides should be nil
	if p.GetAllSlides() != nil {
		t.Error("slides should be nil after close")
	}
}


// =============================================================================
// Test 41: WriteTo with nil presentation
// =============================================================================
func TestComprehensiveWriteNilPresentation(t *testing.T) {
	w := &PPTXWriter{presentation: nil}
	var buf bytes.Buffer
	if err := w.WriteTo(&buf); err == nil {
		t.Error("expected error for nil presentation")
	}
}

// =============================================================================
// Test 42: Multiple charts on same slide
// =============================================================================
func TestComprehensiveMultipleCharts(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// Chart 1: Bar
	c1 := slide.CreateChartShape()
	c1.SetPosition(Inch(0.5), Inch(0.5))
	c1.SetSize(Inch(4), Inch(3))
	c1.GetTitle().SetText("Bar Chart")
	bar := NewBarChart()
	bar.AddSeries(NewChartSeriesOrdered("Data", []string{"A", "B"}, []float64{10, 20}))
	c1.GetPlotArea().SetType(bar)

	// Chart 2: Pie
	c2 := slide.CreateChartShape()
	c2.SetPosition(Inch(5), Inch(0.5))
	c2.SetSize(Inch(4), Inch(3))
	c2.GetTitle().SetText("Pie Chart")
	pie := NewPieChart()
	pie.AddSeries(NewChartSeriesOrdered("Data", []string{"X", "Y"}, []float64{60, 40}))
	c2.GetPlotArea().SetType(pie)

	var buf bytes.Buffer
	if err := p.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo failed: %v", err)
	}

	data := buf.Bytes()
	r := bytes.NewReader(data)
	_, err := ReadFrom(r, int64(len(data)))
	if err != nil {
		t.Fatalf("ReadFrom failed: %v", err)
	}
}

// =============================================================================
// Test 43: Image + chart + text on same slide
// =============================================================================
func TestComprehensiveMixedContent(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	// Text
	rt := slide.CreateRichTextShape()
	rt.SetPosition(Inch(0.5), Inch(0.5))
	rt.SetSize(Inch(4), Inch(1))
	rt.CreateTextRun("Mixed content slide")

	// Image
	img := slide.CreateDrawingShape()
	img.SetImageData(testPNG(), "image/png")
	img.SetPosition(Inch(0.5), Inch(2))
	img.SetSize(Inch(3), Inch(3))

	// Chart
	chart := slide.CreateChartShape()
	chart.SetPosition(Inch(5), Inch(2))
	chart.SetSize(Inch(4), Inch(3))
	chart.GetTitle().SetText("Mixed Chart")
	line := NewLineChart()
	line.AddSeries(NewChartSeriesOrdered("Trend", []string{"1", "2", "3"}, []float64{5, 15, 10}))
	chart.GetPlotArea().SetType(line)

	// Table
	tbl := slide.CreateTableShape(2, 2)
	tbl.SetPosition(Inch(0.5), Inch(6))
	tbl.SetSize(Inch(4), Inch(1))
	tbl.GetCell(0, 0).SetText("A")
	tbl.GetCell(0, 1).SetText("B")
	tbl.GetCell(1, 0).SetText("C")
	tbl.GetCell(1, 1).SetText("D")

	pres := roundTrip(t, p)
	rs, _ := pres.GetSlide(0)

	var hasText, hasImage, hasTable bool
	for _, sh := range rs.GetShapes() {
		switch sh.(type) {
		case *RichTextShape:
			hasText = true
		case *DrawingShape:
			hasImage = true
		case *TableShape:
			hasTable = true
		}
	}

	if !hasText {
		t.Error("expected text shape")
	}
	if !hasImage {
		t.Error("expected image shape")
	}
	if !hasTable {
		t.Error("expected table shape")
	}
}

// =============================================================================
// Test 44: Presentation-level text extraction with notes
// =============================================================================
func TestComprehensiveTextExtractionWithNotes(t *testing.T) {
	p := New()
	s1 := p.GetActiveSlide()
	s1.SetNotes("Slide 1 notes")
	rt := s1.CreateRichTextShape()
	rt.SetPosition(0, 0)
	rt.SetSize(Inch(1), Inch(1))
	rt.CreateTextRun("Slide 1 text")

	s2 := p.CreateSlide()
	s2.SetNotes("Slide 2 notes")
	rt2 := s2.CreateRichTextShape()
	rt2.SetPosition(0, 0)
	rt2.SetSize(Inch(1), Inch(1))
	rt2.CreateTextRun("Slide 2 text")

	text := p.ExtractText()
	for _, expected := range []string{"Slide 1 text", "Slide 1 notes", "Slide 2 text", "Slide 2 notes"} {
		if !strings.Contains(text, expected) {
			t.Errorf("expected '%s' in extracted text", expected)
		}
	}
}

// =============================================================================
// Test 45: Scatter chart with smooth lines
// =============================================================================
func TestComprehensiveScatterChart(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.SetPosition(Inch(1), Inch(1))
	chart.SetSize(Inch(6), Inch(4))
	chart.GetTitle().SetText("Scatter Plot")

	scatter := NewScatterChart()
	scatter.SetSmooth(true)
	scatter.AddSeries(NewChartSeriesOrdered("Points", []string{"1", "2", "3", "4", "5"}, []float64{2, 4, 1, 5, 3}))
	chart.GetPlotArea().SetType(scatter)

	var buf bytes.Buffer
	if err := p.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo failed: %v", err)
	}
}

// =============================================================================
// Test 46: Doughnut chart
// =============================================================================
func TestComprehensiveDoughnutChart(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.SetPosition(Inch(1), Inch(1))
	chart.SetSize(Inch(6), Inch(4))
	chart.GetTitle().SetText("Doughnut")

	doughnut := NewDoughnutChart()
	doughnut.AddSeries(NewChartSeriesOrdered("Parts", []string{"A", "B", "C"}, []float64{30, 50, 20}))
	chart.GetPlotArea().SetType(doughnut)

	var buf bytes.Buffer
	if err := p.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo failed: %v", err)
	}
}

// =============================================================================
// Test 47: Radar chart
// =============================================================================
func TestComprehensiveRadarChart(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.SetPosition(Inch(1), Inch(1))
	chart.SetSize(Inch(6), Inch(4))
	chart.GetTitle().SetText("Radar")

	radar := NewRadarChart()
	radar.AddSeries(NewChartSeriesOrdered("Skills", []string{"Go", "Python", "JS", "Rust", "C++"}, []float64{9, 7, 8, 6, 5}))
	chart.GetPlotArea().SetType(radar)

	var buf bytes.Buffer
	if err := p.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo failed: %v", err)
	}
}

// =============================================================================
// Test 48: Area chart
// =============================================================================
func TestComprehensiveAreaChart(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	chart := slide.CreateChartShape()
	chart.SetPosition(Inch(1), Inch(1))
	chart.SetSize(Inch(6), Inch(4))
	chart.GetTitle().SetText("Area")

	area := NewAreaChart()
	area.AddSeries(NewChartSeriesOrdered("Growth", []string{"2020", "2021", "2022", "2023"}, []float64{10, 25, 40, 55}))
	chart.GetPlotArea().SetType(area)

	var buf bytes.Buffer
	if err := p.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo failed: %v", err)
	}
}

