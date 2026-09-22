package gopresentation

import (
	"strings"
	"testing"
)

// A wrapped line must never open with whitespace: the line break consumes
// what it broke on. Real decks carry double spaces ("Project  declared"), and
// rendering the leftover space at the start of the continuation line showed up
// as an indent PowerPoint does not draw (slide04 of the comparison deck).
func TestWrapLinesNeverStartWithWhitespace(t *testing.T) {
	text := "aaaa bbbb  cccc dddd"
	for w := 40; w <= 300; w++ {
		lines := wrapRunsForTest([]textRun{asciiRun(text)}, w)
		for i, ln := range lines {
			if strings.HasPrefix(ln, " ") || strings.HasPrefix(ln, "\t") {
				t.Errorf("width %d: line %d %q starts with whitespace", w, i, ln)
			}
			if strings.TrimSpace(ln) == "" {
				t.Errorf("width %d: line %d is blank", w, i)
			}
		}
		// Nothing but separators may be dropped by the wrapping.
		joined := strings.Join(lines, " ")
		for _, word := range strings.Fields(text) {
			if !strings.Contains(joined, word) {
				t.Errorf("width %d: word %q lost by wrapping (lines %q)", w, word, lines)
			}
		}
	}
}

// A single space at the break is consumed the same way.
func TestWrapBreakConsumesSingleSpace(t *testing.T) {
	lines := wrapRunsForTest([]textRun{asciiRun("hello world world")}, 80)
	for i, ln := range lines {
		if strings.HasPrefix(ln, " ") {
			t.Errorf("line %d %q starts with whitespace", i, ln)
		}
	}
}
