package gopresentation

import (
	"strings"
	"testing"
)

// Speaker notes: a second text body, written and read by hand.
//
// A note is a text body like any other — one <a:p> per paragraph, a run inside
// each — but the writer put the whole note in one run and the reader joined
// every run it saw with no separator at all. PowerPoint writes a note typed over
// three lines as three paragraphs, so reading a real deck collapsed it into one
// line of text, and opening and saving flattened it for good. Nothing tested any
// of this.

// notesPart builds the markup PowerPoint writes for a notes slide: a body
// placeholder holding the note, and a slide-number placeholder whose text is a
// field rather than something the speaker typed.
func notesPart(body string, extraPlaceholder bool) []byte {
	slideNum := ""
	if extraPlaceholder {
		slideNum = `
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Slide Number Placeholder"/>
          <p:nvPr>
            <p:ph type="sldNum" idx="10"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:fld id="{2A2E4B7C-1111-2222-3333-444455556666}" type="slidenum">
              <a:rPr lang="en-US"/>
              <a:t>7</a:t>
            </a:fld>
          </a:p>
        </p:txBody>
      </p:sp>`
	}

	return []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Notes Placeholder"/>
          <p:nvPr>
            <p:ph type="body" idx="1"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
` + body + `        </p:txBody>
      </p:sp>` + slideNum + `
    </p:spTree>
  </p:cSld>
</p:notes>`)
}

// notesParagraph renders one paragraph of a notes body holding a single run.
func notesParagraph(text string) string {
	return `          <a:p>
            <a:r>
              <a:rPr lang="en-US" dirty="0"/>
              <a:t>` + text + `</a:t>
            </a:r>
          </a:p>
`
}

// TestNotesKeepParagraphBreaks reads the shape PowerPoint writes: one paragraph
// per line, a line break inside a paragraph, a paragraph of its own that holds
// nothing, and a slide-number placeholder that must not be mistaken for notes.
func TestNotesKeepParagraphBreaks(t *testing.T) {
	body := `          <a:p>
            <a:r>
              <a:rPr lang="en-US" dirty="0"/>
              <a:t>first line</a:t>
            </a:r>
            <a:br/>
            <a:r>
              <a:rPr lang="en-US" dirty="0"/>
              <a:t>after a break</a:t>
            </a:r>
          </a:p>
` + notesParagraph("second paragraph") + "          <a:p/>\n"

	want := "first line\nafter a break\nsecond paragraph\n"
	got := (&PPTXReader{}).parseNotesXML(notesPart(body, true))
	if got != want {
		t.Errorf("notes = %q, want %q", got, want)
	}
}

// TestNotesIgnoreAPlaceholderThatHoldsNoRun is the same part without the
// slide-number placeholder's field being read: a text body whose paragraphs hold
// no run is not notes, it is the number the notes page prints.
func TestNotesIgnoreAPlaceholderThatHoldsNoRun(t *testing.T) {
	part := notesPart(notesParagraph("only real notes"), true)
	if got := (&PPTXReader{}).parseNotesXML(part); got != "only real notes" {
		t.Errorf("notes = %q, want %q: a placeholder's body was read as notes", got, "only real notes")
	}

	without := notesPart(notesParagraph("only real notes"), false)
	if got := (&PPTXReader{}).parseNotesXML(without); got != "only real notes" {
		t.Errorf("notes = %q, want %q", got, "only real notes")
	}
}

// TestNotesSurviveRoundTrip is the whole path: the line structure has to reach
// the part as paragraphs and come back as the same string.
func TestNotesSurviveRoundTrip(t *testing.T) {
	const notes = "First line\nsecond line\n\nthird line"

	p := New()
	p.GetActiveSlide().SetNotes(notes)
	pres := roundTrip(t, p)

	if got := pres.GetAllSlides()[0].GetNotes(); got != notes {
		t.Errorf("notes after round trip = %q, want %q", got, notes)
	}

	part := string(zipParts(t, writeToBytes(t, p))["ppt/notesSlides/notesSlide1.xml"])
	if n := strings.Count(part, "<a:p>") + strings.Count(part, "<a:p/>"); n != 4 {
		t.Errorf("the notes part has %d paragraphs, want 4 (one per line)\n%s", n, part)
	}
}
