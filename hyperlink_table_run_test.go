package gopresentation

import (
	"bytes"
	"encoding/xml"
	"strings"
	"testing"
)

// Text runs outside the main emitter.
//
// A run is written by one emitter and read by one parser, and both have to
// cover every place a run can appear. Two places did not:
//
//   - A table cell built its own <a:r> by hand. It kept the size and dropped
//     everything else the font carried — name, East Asian name, weight, slant,
//     underline, strikethrough, colour and the hyperlink. The reader reads all
//     of it back, so a load and a save flattened a formatted table.
//   - Hyperlinks were write-only and half-written: the reader never parsed
//     <a:hlinkClick> at all, so an open and a save destroyed every link in a
//     deck, and an internal link — the documented NewInternalHyperlink — was
//     never written even once, because the writer skipped internal links and
//     the relationship was only ever external.
//
// The assertions are structural where the bug is structural: a relationship id
// in a part that no relationship defines is what makes PowerPoint ask to repair
// the file, and a round trip alone would not see it.

// formattedTableRun returns a presentation whose only shape is a one-cell table
// holding one run with every property a run can carry.
func formattedTableRun(t *testing.T) (*Presentation, *TextRun) {
	t.Helper()
	p := New()
	table := NewTableShape(1, 1)
	table.SetOffsetX(500000)
	table.SetOffsetY(500000)
	table.SetWidth(6000000)
	table.SetHeight(1200000)

	cell := table.GetCell(0, 0)
	cell.SetText("TABLE-CELL-TEXT")
	run := cell.GetParagraphs()[0].elements[0].(*TextRun)
	run.GetFont().
		SetName("Segoe UI").
		SetSize(24).
		SetBold(true).
		SetItalic(true).
		SetUnderline(UnderlineSingle).
		SetStrikethrough(true).
		SetColor(ColorRed)
	run.GetFont().NameEA = "Microsoft YaHei"

	p.GetActiveSlide().AddShape(table)
	return p, run
}

// cellXML returns the first table cell of a slide part, so that an assertion
// about formatting is scoped to the cell and cannot be satisfied by a run
// somewhere else on the slide.
func cellXML(t *testing.T, slide []byte) string {
	t.Helper()
	text := string(slide)
	open := strings.Index(text, "<a:tc>")
	end := strings.Index(text, "</a:tc>")
	if open < 0 || end < open {
		t.Fatalf("no table cell in the slide part:\n%s", text)
	}
	return text[open:end]
}

// TestTableRunFormattingReachesThePackage pins the writer half: everything the
// font carries has to be in the cell's <a:rPr>, not only the size.
func TestTableRunFormattingReachesThePackage(t *testing.T) {
	p, _ := formattedTableRun(t)
	cell := cellXML(t, zipParts(t, writeToBytes(t, p))["ppt/slides/slide1.xml"])

	for _, want := range []string{
		`sz="2400"`,
		`b="1"`,
		`i="1"`,
		`u="sng"`,
		`strike="sngStrike"`,
		`<a:latin typeface="Segoe UI"/>`,
		`<a:ea typeface="Microsoft YaHei"/>`,
		`<a:srgbClr val="FF0000"/>`,
		"TABLE-CELL-TEXT",
	} {
		if !strings.Contains(cell, want) {
			t.Errorf("the table cell's run has no %s:\n%s", want, cell)
		}
	}
}

// TestTableRunFormattingSurvivesRoundTrip pins the consequence: the reader has
// always read these properties, so before the writer emitted them a load and a
// save silently flattened the cell's text.
func TestTableRunFormattingSurvivesRoundTrip(t *testing.T) {
	p, _ := formattedTableRun(t)
	run := firstTableRun(t, roundTrip(t, p))
	if run == nil {
		t.Fatal("the round trip lost the table or its text")
	}

	font := run.GetFont()
	if font == nil {
		t.Fatal("the round-tripped run has no font")
	}
	if !font.Bold || !font.Italic || !font.Strikethrough {
		t.Errorf("bold=%v italic=%v strike=%v, want all true",
			font.Bold, font.Italic, font.Strikethrough)
	}
	if font.Underline != UnderlineSingle {
		t.Errorf("underline = %q, want %q", font.Underline, UnderlineSingle)
	}
	if font.Name != "Segoe UI" {
		t.Errorf("font name = %q, want %q", font.Name, "Segoe UI")
	}
	if font.NameEA != "Microsoft YaHei" {
		t.Errorf("East Asian font name = %q, want %q", font.NameEA, "Microsoft YaHei")
	}
	if font.Size != 24 {
		t.Errorf("font size = %d, want 24", font.Size)
	}
	if font.Color.ARGB != ColorRed.ARGB {
		t.Errorf("colour = %q, want %q", font.Color.ARGB, ColorRed.ARGB)
	}
}

// firstTableRun returns the first text run inside a table of the presentation.
func firstTableRun(t *testing.T, pres *Presentation) *TextRun {
	t.Helper()
	for _, slide := range pres.GetAllSlides() {
		for _, shape := range flattenShapes(slide.GetShapes()) {
			table, ok := shape.(*TableShape)
			if !ok {
				continue
			}
			for _, row := range table.GetRows() {
				for _, cell := range row {
					for _, para := range cell.GetParagraphs() {
						for _, elem := range para.elements {
							if tr, ok := elem.(*TextRun); ok {
								return tr
							}
						}
					}
				}
			}
		}
	}
	return nil
}

// hlinkClickRef is one <a:hlinkClick> as it was written.
type hlinkClickRef struct {
	id     string
	action string
}

// hlinkClicks returns every <a:hlinkClick> in a part. They are read back out of
// the written file rather than assumed, because a link whose r:id names no
// relationship is the failure that matters.
func hlinkClicks(t *testing.T, part []byte) []hlinkClickRef {
	t.Helper()
	var out []hlinkClickRef
	dec := xml.NewDecoder(bytes.NewReader(part))
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "hlinkClick" {
			continue
		}
		var ref hlinkClickRef
		for _, attr := range start.Attr {
			switch {
			case attr.Name.Space == nsOfficeDocRels && attr.Name.Local == "id":
				ref.id = attr.Value
			case attr.Name.Space == "" && attr.Name.Local == "action":
				ref.action = attr.Value
			}
		}
		out = append(out, ref)
	}
	return out
}

// TestInternalHyperlinkIsWrittenAndResolves covers the link kind the writer used
// to skip outright. A jump to another slide is an internal relationship of slide
// type plus the jump action; neither alone is a link, so both are asserted, and
// the id in the part is looked up in the relationship file it names.
func TestInternalHyperlinkIsWrittenAndResolves(t *testing.T) {
	p := New()
	run := p.GetActiveSlide().CreateRichTextShape().CreateTextRun("go to slide 2")
	run.SetHyperlink(NewInternalHyperlink(2))
	p.CreateSlide()

	pkg := writeToBytes(t, p)
	parts := zipParts(t, pkg)

	const part = "ppt/slides/slide1.xml"
	clicks := hlinkClicks(t, parts[part])
	if len(clicks) != 1 {
		t.Fatalf("slide 1 has %d hlinkClick elements, want 1: the internal hyperlink was not written", len(clicks))
	}
	if clicks[0].action != actionSlideJump {
		t.Errorf("action = %q, want %q", clicks[0].action, actionSlideJump)
	}

	rels, err := readPackageRelationships(parts, relsPathFor(part))
	if err != nil {
		t.Fatalf("read the slide relationships: %v", err)
	}
	rel, ok := rels[clicks[0].id]
	if !ok {
		t.Fatalf("the hyperlink refers to %s, which slide1.xml.rels does not define", clicks[0].id)
	}
	if rel.target != "slide2.xml" {
		t.Errorf("hyperlink target = %q, want slide2.xml", rel.target)
	}
}

// TestInternalHyperlinkOutOfRangeLeavesNoReference is the other half: a target
// no slide backs is not writable, and the three places that number these
// relationships have to agree about that. A run that still carried the id
// placeholder, or a relationship pointing at a slide part that does not exist,
// is the damage this refuses to do.
func TestInternalHyperlinkOutOfRangeLeavesNoReference(t *testing.T) {
	p := New() // one slide only
	run := p.GetActiveSlide().CreateRichTextShape().CreateTextRun("nowhere")
	run.SetHyperlink(NewInternalHyperlink(9))

	pkg := writeToBytes(t, p)
	parts := zipParts(t, pkg)

	const part = "ppt/slides/slide1.xml"
	if clicks := hlinkClicks(t, parts[part]); len(clicks) != 0 {
		t.Errorf("a link to slide 9 of a one-slide deck produced %d hlinkClick elements", len(clicks))
	}
	if strings.Contains(string(parts[part]), "rId_hlink_") {
		t.Error("slide 1 still holds an unsubstituted hyperlink id placeholder")
	}

	rels, err := readPackageRelationships(parts, relsPathFor(part))
	if err != nil {
		t.Fatalf("read the slide relationships: %v", err)
	}
	for id, rel := range rels {
		if strings.HasSuffix(rel.target, "slide9.xml") {
			t.Errorf("relationship %s targets slide9.xml, which the package does not contain", id)
		}
	}
}

// TestHyperlinkSurvivesRoundTrip covers the reader half for both kinds of link.
// Neither used to come back: the reader had no case for <a:hlinkClick>, so a
// presentation opened and saved again lost every link it contained.
func TestHyperlinkSurvivesRoundTrip(t *testing.T) {
	p := New()
	external := p.GetActiveSlide().CreateRichTextShape().CreateTextRun("external")
	external.SetHyperlink(NewHyperlink("https://example.com/deep/link"))
	p.CreateSlide() // slide 2, the target of the internal link
	internal := p.CreateSlide().CreateRichTextShape().CreateTextRun("internal")
	internal.SetHyperlink(NewInternalHyperlink(2))

	got := map[string]*Hyperlink{}
	for _, slide := range roundTrip(t, p).GetAllSlides() {
		for _, shape := range flattenShapes(slide.GetShapes()) {
			for _, para := range shapeParagraphs(shape) {
				for _, elem := range para.elements {
					if tr, ok := elem.(*TextRun); ok {
						got[tr.text] = tr.GetHyperlink()
					}
				}
			}
		}
	}

	link := got["external"]
	if link == nil {
		t.Fatal("the external hyperlink was lost by the round trip")
	}
	if link.IsInternal {
		t.Error("the external hyperlink came back as an internal link")
	}
	if link.URL != "https://example.com/deep/link" {
		t.Errorf("external URL = %q, want %q", link.URL, "https://example.com/deep/link")
	}

	link = got["internal"]
	if link == nil {
		t.Fatal("the internal hyperlink was lost by the round trip")
	}
	if !link.IsInternal {
		t.Error("the internal hyperlink came back as an external link")
	}
	if link.SlideNumber != 2 {
		t.Errorf("internal link target = slide %d, want slide 2", link.SlideNumber)
	}
}

// TestTableHyperlinkGetsARelationship is the multi-part trap in one test: the
// cell's run now emits the id placeholder, so the walk that numbers hyperlinks
// has to reach into table cells too. If it does not, the placeholder is never
// substituted and the cell refers to a relationship nothing defines.
func TestTableHyperlinkGetsARelationship(t *testing.T) {
	p := New()
	table := NewTableShape(1, 1)
	table.SetOffsetX(500000)
	table.SetOffsetY(500000)
	table.SetWidth(6000000)
	table.SetHeight(1200000)

	cell := table.GetCell(0, 0)
	cell.SetText("LINKED-CELL")
	run := cell.GetParagraphs()[0].elements[0].(*TextRun)
	run.SetHyperlink(NewHyperlink("https://example.com/cell"))
	p.GetActiveSlide().AddShape(table)

	pkg := writeToBytes(t, p)
	parts := zipParts(t, pkg)

	const part = "ppt/slides/slide1.xml"
	clicks := hlinkClicks(t, parts[part])
	if len(clicks) != 1 {
		t.Fatalf("the table cell has %d hlinkClick elements, want 1", len(clicks))
	}
	if strings.HasPrefix(clicks[0].id, "rId_hlink_") {
		t.Fatalf("the cell's hyperlink kept its placeholder id %q", clicks[0].id)
	}

	rels, err := readPackageRelationships(parts, relsPathFor(part))
	if err != nil {
		t.Fatalf("read the slide relationships: %v", err)
	}
	rel, ok := rels[clicks[0].id]
	if !ok {
		t.Fatalf("the cell's hyperlink refers to %s, which slide1.xml.rels does not define", clicks[0].id)
	}
	if rel.target != "https://example.com/cell" {
		t.Errorf("hyperlink target = %q, want the cell's URL", rel.target)
	}

	run = firstTableRun(t, roundTrip(t, p))
	if run == nil || run.GetHyperlink() == nil {
		t.Fatal("the table cell's hyperlink was lost by the round trip")
	}
	if url := run.GetHyperlink().URL; url != "https://example.com/cell" {
		t.Errorf("URL after round trip = %q, want the cell's URL", url)
	}
}
