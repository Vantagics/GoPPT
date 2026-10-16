package gopresentation

import (
	"strings"
	"testing"
)

// TestExtractTextReachesEveryShapeInsideGroups is a regression guard for a
// silent data loss: Slide.ExtractText had two switches, and the one inside the
// *GroupShape case recognised only RichTextShape and PlaceholderShape. An
// AutoShape or a table inside a group therefore contributed nothing, and
// neither did a group inside a group, while the identical shape at the top
// level did. Every real deck groups things.
//
// It is also the assertion that keeps the SmartArt reader honest: a diagram
// arrives as a group of AutoShapes, so without this the diagram's text would
// be read, drawn, and then invisible to anything asking the model what the
// slide says.
func TestExtractTextReachesEveryShapeInsideGroups(t *testing.T) {
	pres := New()
	slide := newSlide()
	pres.AddSlide(slide)

	rich := NewRichTextShape()
	rich.CreateTextRun("GroupedRich")
	outer := NewGroupShape()
	outer.AddShape(rich)
	outer.AddShape(NewAutoShape().SetText("GroupedAuto"))

	tbl := NewTableShape(1, 1)
	tbl.GetCell(0, 0).SetText("GroupedCell")
	outer.AddShape(tbl)

	inner := NewGroupShape()
	inner.AddShape(NewAutoShape().SetText("NestedAuto"))
	outer.AddShape(inner)

	// The same four strings at the top level, which is the half that already
	// worked: if only the grouped copies came back, the fix has over-corrected.
	slide.AddShape(NewAutoShape().SetText("TopAuto"))
	slide.AddShape(outer)

	got := slide.ExtractText()
	for _, want := range []string{
		"GroupedRich", "GroupedAuto", "GroupedCell", "NestedAuto", "TopAuto",
	} {
		if !strings.Contains(got, want) {
			t.Errorf("ExtractText() = %q, missing %q", got, want)
		}
	}
}
