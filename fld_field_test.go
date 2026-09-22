package gopresentation

import (
	"bytes"
	"strings"
	"testing"
)

// <a:fld> — a field, not a run.
//
// A field wraps run content in a CT_TextField and is evaluated when the deck
// is displayed: the slide-number field's cached <a:t> is whatever the last
// save saw, so the renderer must re-resolve it per slide and the writer must
// emit a field again. Three halves must agree. Before any of this existed the
// reader skipped the field's <a:t> outright (its gate matches on <a:r>), every
// slide-number placeholder rendered empty, and a save dropped the field for
// good.

// fldPackage builds a one-slide package whose only shape is the slide-number
// placeholder holding one field, so the reader's path can be exercised whole.
func fldPackage(t *testing.T, cached string) []byte {
	t.Helper()
	return buildZip(t, map[string][]byte{
		"ppt/presentation.xml": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"` +
			` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
			`<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`),
		"ppt/_rels/presentation.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
			`<Relationship Id="rId1"` +
			` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"` +
			` Target="slides/slide1.xml"/></Relationships>`),
		"ppt/slides/slide1.xml": []byte(bodyPrSlide(`<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Slide Number Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p><a:fld id="{2A2E4B7C-1111-2222-3333-444455556666}" type="slidenum"><a:rPr lang="en-US"/><a:t>` + cached + `</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>
  </p:txBody>
</p:sp>`)),
	})
}

// readSlideNumberPlaceholder returns the sldNum placeholder of the one-slide
// package fldPackage builds.
func readSlideNumberPlaceholder(t *testing.T, cached string) *PlaceholderShape {
	t.Helper()
	reader, err := NewReader(ReaderPowerPoint2007)
	if err != nil {
		t.Fatalf("new reader: %v", err)
	}
	pres, err := reader.ReadFromReader(bytes.NewReader(fldPackage(t, cached)), int64(len(fldPackage(t, cached))))
	if err != nil {
		t.Fatalf("read package: %v", err)
	}
	slide, err := pres.GetSlide(0)
	if err != nil {
		t.Fatalf("get slide: %v", err)
	}
	for _, s := range slide.GetShapes() {
		if ph, ok := s.(*PlaceholderShape); ok && ph.GetPlaceholderType() == PlaceholderSlideNum {
			return ph
		}
	}
	t.Fatal("no slide-number placeholder on the slide")
	return nil
}

// TestSlideNumberFieldReadsAsAFieldRun pins the reader half: the field's <a:t>
// becomes a run carrying the field's type, so neither the text nor the fact
// that it is a field is lost on the way into the model.
func TestSlideNumberFieldReadsAsAFieldRun(t *testing.T) {
	ph := readSlideNumberPlaceholder(t, "7")
	paras := ph.GetParagraphs()
	if len(paras) != 1 {
		t.Fatalf("paragraph count = %d, want 1", len(paras))
	}
	runs := 0
	for _, el := range paras[0].GetElements() {
		tr, ok := el.(*TextRun)
		if !ok {
			continue
		}
		runs++
		if tr.GetText() != "7" {
			t.Errorf("field text = %q, want %q", tr.GetText(), "7")
		}
		if tr.GetFieldType() != "slidenum" {
			t.Errorf("field type = %q, want %q", tr.GetFieldType(), "slidenum")
		}
	}
	if runs != 1 {
		t.Errorf("run count = %d, want 1: the field's <a:t> was read %d times or not at all", runs, runs)
	}
}

// TestSlideNumberFieldEvaluatesPerSlide pins the renderer half: a slidenum
// field draws the number of the slide on the canvas, not the cached text —
// and does so even when the cache is empty, which is how a generated deck
// arrives. A plain run with the same text keeps its text: the substitution
// keys on the field type, not on the digits.
func TestSlideNumberFieldEvaluatesPerSlide(t *testing.T) {
	r := &renderer{slideNumber: 7}

	field := &TextRun{text: "42", font: NewFont()}
	field.SetFieldType("slidenum")
	runs := r.buildParaTextRuns([]ParagraphElement{field})
	if len(runs) != 1 || runs[0].text != "7" {
		t.Errorf("slidenum field rendered %v, want the slide number 7", runTexts(runs))
	}

	empty := &TextRun{font: NewFont()}
	empty.SetFieldType("slidenum")
	runs = r.buildParaTextRuns([]ParagraphElement{empty})
	if len(runs) != 1 || runs[0].text != "7" {
		t.Errorf("empty-cache field rendered %v, want 7 anyway", runTexts(runs))
	}

	plain := &TextRun{text: "42", font: NewFont()}
	runs = r.buildParaTextRuns([]ParagraphElement{plain})
	if len(runs) != 1 || runs[0].text != "42" {
		t.Errorf("plain run rendered %v, want its own text 42", runTexts(runs))
	}
}

// runTexts joins the rendered text of each run, for error messages.
func runTexts(runs []textRun) []string {
	out := make([]string, len(runs))
	for i, r := range runs {
		out[i] = r.text
	}
	return out
}

// TestWriterKeepsSlideNumberFieldLive pins the writer half: a run marked as a
// field goes back as <a:fld type="slidenum">, not as a literal <a:r> — a
// flattened field freezes its cached number into the file, and a deck saved
// after reordering shows stale numbers forever after. A plain run's XML must
// stay fieldless.
func TestWriterKeepsSlideNumberFieldLive(t *testing.T) {
	p := New()
	shape := p.GetActiveSlide().CreateRichTextShape()
	field := shape.CreateTextRun("12")
	field.SetFieldType("slidenum")

	part := string(zipParts(t, writeToBytes(t, p))["ppt/slides/slide1.xml"])
	if !strings.Contains(part, `<a:fld id="`) || !strings.Contains(part, `type="slidenum"`) {
		t.Errorf("field run written as %q, want an <a:fld type=\"slidenum\"> wrapper", part)
	}
	// The cached text still travels, as the field's content.
	if !strings.Contains(part, "<a:t>12</a:t>") {
		t.Errorf("field text missing from %q", part)
	}

	p2 := New()
	shape2 := p2.GetActiveSlide().CreateRichTextShape()
	shape2.CreateTextRun("12")
	part2 := string(zipParts(t, writeToBytes(t, p2))["ppt/slides/slide1.xml"])
	if strings.Contains(part2, "<a:fld") {
		t.Errorf("plain run written as %q, want no field wrapper", part2)
	}
}

// TestSlideNumberPlaceholderSkipsTxStyles pins the ladder decision the master's
// sldNum placeholder forced: sldNum/dt/ftr are chrome placeholders that
// PowerPoint styles only from their own placeholder <a:lstStyle>, never from
// <p:txStyles>. Falling through to otherStyle gave the page number a 24pt
// bullet list — a bullet glyph and a left-aligned, oversized number exactly
// where PowerPoint draws none of that.
func TestSlideNumberPlaceholderSkipsTxStyles(t *testing.T) {
	ph := NewPlaceholderShape(PlaceholderSlideNum)
	para := ph.CreateParagraph()
	tr := para.CreateTextRun("7")
	tr.font = NewFont()
	tr.font.Size = 18

	other := &masterLevelStyle{
		size:   24,
		align:  "l",
		marL:   228600,
		indent: -228600,
		buChar: "•",
		buFont: "Arial",
	}
	m := &masterTextStyles{}
	m.other[1] = other

	applyMasterTextStyles(ph, m)

	if para.bullet != nil {
		t.Errorf("sldNum paragraph took a bullet %+v from txStyles; PowerPoint styles chrome placeholders from their own lstStyle only", para.bullet)
	}
	if tr.font.Size != 18 {
		t.Errorf("sldNum run size = %d, want the untouched 18 (the placeholder's own defRPr 12 overrides instead)", tr.font.Size)
	}
	if para.alignment != nil && para.alignment.MarginLeft != 0 {
		t.Errorf("sldNum paragraph took marL %d from txStyles", para.alignment.MarginLeft)
	}
}

// TestPlaceholderLstStyleAlignsParagraphs pins the alignment rung the same
// sldNum case needs: the master's slide-number placeholder says algn="r" in
// its lstStyle lvl1pPr, and that is the only place the right alignment comes
// from. An explicit paragraph alignment keeps its own; a definition that says
// nothing aligns nothing.
func TestPlaceholderLstStyleAlignsParagraphs(t *testing.T) {
	ph := NewPlaceholderShape(PlaceholderSlideNum)
	para := ph.CreateParagraph()
	para.CreateTextRun("7").font = NewFont()

	def := layoutPlaceholder{alignH: "r", alignSet: true}
	applyPlaceholderAlign(ph, &def)
	if para.alignment == nil || para.alignment.Horizontal != HorizontalAlignment("r") {
		t.Errorf("paragraph alignment = %+v, want r from the placeholder lstStyle", para.alignment)
	}

	// An explicit alignment wins over the inherited one.
	ph2 := NewPlaceholderShape(PlaceholderSlideNum)
	para2 := ph2.CreateParagraph()
	para2.CreateTextRun("7").font = NewFont()
	para2.alignment = NewAlignment()
	para2.alignment.Horizontal = HorizontalAlignment("l")
	applyPlaceholderAlign(ph2, &def)
	if para2.alignment.Horizontal != HorizontalAlignment("l") {
		t.Errorf("explicit alignment overwritten: %+v", para2.alignment)
	}

	// A definition that declares no alignment aligns nothing.
	ph3 := NewPlaceholderShape(PlaceholderSlideNum)
	para3 := ph3.CreateParagraph()
	para3.CreateTextRun("7").font = NewFont()
	applyPlaceholderAlign(ph3, &layoutPlaceholder{})
	if para3.alignment.Horizontal != "" {
		t.Errorf("alignment invented from a silent definition: %+v", para3.alignment)
	}
}
