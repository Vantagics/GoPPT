package gopresentation

import (
	"strings"
	"testing"
	"time"
)

// Comments, end to end.
//
// A comment is split across two parts: the text body lives in
// ppt/comments/commentN.xml, but the author's identity lives in a separate
// package part, ppt/commentAuthors.xml, that the comment only references by id.
// Both halves therefore have to be written and read together.
//
// Every defect covered here has the same shape as the chart-styling bugs: the
// round trip looked correct because both sides agreed on the same lossy answer.
// The author name was never read back and the writer groups comments into
// authors by name, so on the next save every comment collapsed into a single
// author with an empty name. Nothing failed — the information was just gone.

// TestCommentAuthorsSurviveRoundTrip covers the author table. The writer emits
// ppt/commentAuthors.xml and the reader used to ignore it entirely, rebuilding
// each author as a bare {ID} — which also meant every author merged into one on
// the following save, because the writer keys authors by name.
func TestCommentAuthorsSurviveRoundTrip(t *testing.T) {
	p := New()
	slide := p.GetActiveSlide()

	slide.AddComment(NewComment().
		SetAuthor(NewCommentAuthor("Alice Zhang", "AZ")).
		SetText("First note").
		SetPosition(10, 20))
	slide.AddComment(NewComment().
		SetAuthor(NewCommentAuthor("Bob Li", "BL")).
		SetText("Second note").
		SetPosition(30, 40))

	part, ok := zipParts(t, writeToBytes(t, p))["ppt/commentAuthors.xml"]
	if !ok {
		t.Fatal("ppt/commentAuthors.xml was not written")
	}
	if !strings.Contains(string(part), `name="Alice Zhang"`) ||
		!strings.Contains(string(part), `initials="BL"`) {
		t.Errorf("author part does not carry the names; got:\n%s", part)
	}

	got := roundTrip(t, p).GetAllSlides()[0].GetComments()
	if len(got) != 2 {
		t.Fatalf("round trip produced %d comments, want 2", len(got))
	}
	if got[0].Author == nil || got[0].Author.Name != "Alice Zhang" || got[0].Author.Initials != "AZ" {
		t.Errorf("first comment author = %+v, want Alice Zhang/AZ", got[0].Author)
	}
	if got[1].Author == nil || got[1].Author.Name != "Bob Li" || got[1].Author.Initials != "BL" {
		t.Errorf("second comment author = %+v, want Bob Li/BL", got[1].Author)
	}
	// The second author's colour index is only right if clrIdx was actually
	// read; a reader that ignores it leaves both authors at 0.
	if got[1].Author.ColorIdx != 1 {
		t.Errorf("second comment colour index = %d, want 1", got[1].Author.ColorIdx)
	}
}

// TestCommentDateSurvivesRoundTrip covers the timestamp. The writer has always
// emitted dt="…", but the reader ignored the attribute, so a reopened comment
// was stamped with the moment it was read rather than the moment it was written.
func TestCommentDateSurvivesRoundTrip(t *testing.T) {
	p := New()
	p.GetActiveSlide().AddComment(NewComment().
		SetAuthor(NewCommentAuthor("A", "A")).
		SetText("note").
		SetDate(time.Date(2024, 3, 15, 9, 30, 45, 123_000_000, time.UTC)))

	got := roundTrip(t, p).GetAllSlides()[0].GetComments()
	if len(got) != 1 {
		t.Fatalf("round trip produced %d comments, want 1", len(got))
	}
	want := time.Date(2024, 3, 15, 9, 30, 45, 123_000_000, time.UTC)
	if !got[0].Date.Equal(want) {
		t.Errorf("comment date = %s, want %s", got[0].Date.UTC(), want)
	}
}

// TestCommentTextIsWrittenAsATextBody is a structural check on the emitted XML.
// p:text is a CT_TextBody, so bare character data directly inside it is not
// something PowerPoint will read — the comment would show up empty there even
// though our own reader accepted it. Only inspecting the part text catches this,
// since the round trip is self-consistent either way.
func TestCommentTextIsWrittenAsATextBody(t *testing.T) {
	p := New()
	p.GetActiveSlide().AddComment(NewComment().
		SetAuthor(NewCommentAuthor("A", "A")).
		SetText("Hello & goodbye"))

	part, ok := zipParts(t, writeToBytes(t, p))["ppt/comments/comment1.xml"]
	if !ok {
		t.Fatal("ppt/comments/comment1.xml was not written")
	}
	text := string(part)

	if !strings.Contains(text, "<a:bodyPr/>") || !strings.Contains(text, "<a:p>") {
		t.Errorf("p:text is missing its text-body skeleton; got:\n%s", text)
	}
	if !strings.Contains(text, "<a:t>Hello &amp; goodbye</a:t>") {
		t.Errorf("comment text is not an escaped a:t run; got:\n%s", text)
	}
	if strings.Contains(text, "<p:text>Hello") {
		t.Errorf("comment text is bare character data inside p:text, which PowerPoint does not read; got:\n%s", text)
	}
	if !strings.Contains(text, "xmlns:a=") {
		t.Errorf("the a: prefix is not declared, so the text body is malformed; got:\n%s", text)
	}
}

// TestCommentTextIsReadFromRuns covers the parse side, using the shape
// PowerPoint itself writes: an indented p:text body whose text is in a:t runs.
//
// Reading any character data inside p:text would take the indentation that
// precedes </p:text> as the comment text, because it arrives after the runs and
// the old code assigned rather than accumulated. Runs are also joined, so a
// comment split across several runs is not truncated to its last one.
func TestCommentTextIsReadFromRuns(t *testing.T) {
	const part = `<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cm authorId="7" dt="2024-03-15T09:30:45.123" idx="1">
    <p:pos x="10" y="20"/>
    <p:text>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p>
        <a:r>
          <a:rPr lang="en-US" dirty="0"/>
          <a:t>Hello </a:t>
        </a:r>
        <a:r>
          <a:rPr lang="en-US" dirty="0"/>
          <a:t>world</a:t>
        </a:r>
      </a:p>
      <a:p>
        <a:r>
          <a:rPr lang="en-US" dirty="0"/>
          <a:t>Second line</a:t>
        </a:r>
      </a:p>
    </p:text>
  </p:cm>
</p:cmLst>`

	slide := newSlide()
	authors := map[int]*CommentAuthor{7: {ID: 7, Name: "Alice", Initials: "AZ", ColorIdx: 3}}
	(&PPTXReader{}).parseCommentsXML([]byte(part), slide, authors)

	got := slide.GetComments()
	if len(got) != 1 {
		t.Fatalf("parsed %d comments, want 1", len(got))
	}
	if want := "Hello world\nSecond line"; got[0].Text != want {
		t.Errorf("comment text = %q, want %q", got[0].Text, want)
	}
	if got[0].Author == nil || got[0].Author.Name != "Alice" {
		t.Errorf("author = %+v, want the resolved Alice", got[0].Author)
	}
	if got[0].PositionX != 10 || got[0].PositionY != 20 {
		t.Errorf("position = (%d,%d), want (10,20)", got[0].PositionX, got[0].PositionY)
	}
	if want := time.Date(2024, 3, 15, 9, 30, 45, 123_000_000, time.UTC); !got[0].Date.Equal(want) {
		t.Errorf("date = %s, want %s", got[0].Date.UTC(), want)
	}
}

// TestCommentLegacyTextFormIsStillRead keeps the reader tolerant of the bare
// p:text form this library emitted before the body was fixed. Files written by
// an older version must not start reading back blank.
func TestCommentLegacyTextFormIsStillRead(t *testing.T) {
	const part = `<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cm authorId="0" dt="2024-03-15T09:30:45.123" idx="1">
    <p:pos x="1" y="2"/>
    <p:text>legacy note</p:text>
  </p:cm>
</p:cmLst>`

	slide := newSlide()
	(&PPTXReader{}).parseCommentsXML([]byte(part), slide, nil)

	got := slide.GetComments()
	if len(got) != 1 {
		t.Fatalf("parsed %d comments, want 1", len(got))
	}
	if got[0].Text != "legacy note" {
		t.Errorf("legacy comment text = %q, want %q", got[0].Text, "legacy note")
	}
	// No author table in the part or the lookup, so the id stands alone.
	if got[0].Author == nil || got[0].Author.ID != 0 || got[0].Author.Name != "" {
		t.Errorf("legacy comment author = %+v, want a bare id 0", got[0].Author)
	}
}
