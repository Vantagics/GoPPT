package main

import (
	"fmt"
	"strconv"
	"strings"

	ppt "github.com/Vantagics/GoPPT"
)

// probeRuns prints the text runs the reader built, with the font properties it
// resolved for each.
//
//	go run ./cmd/debug_preview --runs <file.pptx> [slide-number]
//
// The point is to compare what the reader decided against what the XML says.
// A run whose colour reads as text or a fill whose alpha disappeared is then
// visible without rendering anything — and a colour that looks wrong on screen
// can be attributed to the reader rather than guessed at in the renderer.
func probeRuns(path string, only int) {
	reader, err := ppt.NewReader(ppt.ReaderPowerPoint2007)
	if err != nil {
		fmt.Printf("new reader: %v\n", err)
		return
	}
	pres, err := reader.Read(path)
	if err != nil {
		fmt.Printf("read: %v\n", err)
		return
	}
	for i, slide := range pres.GetAllSlides() {
		if only > 0 && i+1 != only {
			continue
		}
		var walk func(shapes []ppt.Shape, depth int)
		walk = func(shapes []ppt.Shape, depth int) {
			for _, s := range shapes {
				if g, ok := s.(*ppt.GroupShape); ok {
					walk(g.GetShapes(), depth+1)
					continue
				}
				rt, ok := s.(runCarrier)
				if !ok {
					continue
				}
				fmt.Printf("slide%02d %s%s box=(%d,%d %dx%d)\n", i+1,
					strings.Repeat("  ", depth), shapeLabel(s),
					rt.GetOffsetX(), rt.GetOffsetY(), rt.GetWidth(), rt.GetHeight())
				for pi, p := range rt.GetParagraphs() {
					// Paragraph properties first: they decide where the runs
					// land, and a run list alone cannot show a wrapping or a
					// hanging-indent bug.
					if al := p.GetAlignment(); al != nil {
						fmt.Printf("  p%d algn=%q lvl=%d marL=%d marR=%d indent=%d\n",
							pi, al.Horizontal, al.Level, al.MarginLeft, al.MarginRight, al.Indent)
					} else {
						fmt.Printf("  p%d algn=<nil>\n", pi)
					}
					for _, e := range p.GetElements() {
						tr, ok := e.(*ppt.TextRun)
						if !ok {
							fmt.Printf("  p%d <%s>\n", pi, e.GetElementType())
							continue
						}
						f := tr.GetFont()
						if f == nil {
							fmt.Printf("  p%d run %s (no font)\n", pi, quoteASCII(tr.GetText()))
							continue
						}
						fmt.Printf("  p%d run %s sz=%d bold=%v italic=%v latin=%q ea=%q color=%s\n",
							pi, quoteASCII(tr.GetText()), f.Size, f.Bold, f.Italic, f.Name, f.NameEA, describeARGB(f.Color.ARGB))
					}
				}
			}
		}
		walk(slide.GetShapes(), 0)
	}
}

// runCarrier is the part of a shape this probe needs. RichTextShape has it, and
// so does every shape that embeds one.
//
// Asking for the interface rather than *RichTextShape matters: PlaceholderShape
// embeds RichTextShape, and a type assertion to the concrete embedded type does
// not match a pointer to the outer one. Asserting the concrete type silently
// skipped every placeholder in the deck — which is where the inherited text
// actually lives, so the probe was blind to the very shapes it was written to
// explain.
type runCarrier interface {
	GetParagraphs() []*ppt.Paragraph
	GetOffsetX() int64
	GetOffsetY() int64
	GetWidth() int64
	GetHeight() int64
}

// shapeLabel names the shape so a report can be read against the XML.
func shapeLabel(s ppt.Shape) string {
	if ph, ok := s.(*ppt.PlaceholderShape); ok {
		return fmt.Sprintf("ph:%-7s ", ph.GetPlaceholderType())
	}
	switch s.GetType() {
	case ppt.ShapeTypeRichText:
		return "textbox  "
	case ppt.ShapeTypeTable:
		return "table    "
	case ppt.ShapeTypeChart:
		return "chart    "
	case ppt.ShapeTypeAutoShape:
		return "autoshape"
	}
	return fmt.Sprintf("%-9s", s.GetType())
}

// quoteASCII renders text as an ASCII-only quoted string, escaping every rune
// outside printable ASCII as \uXXXX.
//
// The host console decodes this program's output with the system codepage, so
// printing CJK directly produces mojibake that cannot be read back. Escaping
// keeps the report legible and, more usefully, states the exact code point —
// which is what matters when the question is "is this ① (U+2460) or something
// else the font happens to have".
func quoteASCII(s string) string {
	var b strings.Builder
	b.WriteByte('"')
	for _, r := range s {
		switch {
		case r == '"' || r == '\\':
			b.WriteByte('\\')
			b.WriteRune(r)
		case r >= 0x20 && r <= 0x7E:
			b.WriteRune(r)
		case r == '\n':
			b.WriteString(`\n`)
		case r == '\t':
			b.WriteString(`\t`)
		default:
			fmt.Fprintf(&b, "\\u%04X", r)
		}
	}
	b.WriteByte('"')
	return b.String()
}

// describeARGB renders an ARGB string with its alpha broken out, because the
// interesting question about a colour is always whether its alpha survived.
func describeARGB(argb string) string {
	if len(argb) != 8 {
		return fmt.Sprintf("%q (malformed)", argb)
	}
	a, err := strconv.ParseUint(argb[:2], 16, 8)
	if err != nil {
		return fmt.Sprintf("%q (malformed)", argb)
	}
	pct := float64(a) * 100 / 255
	return fmt.Sprintf("#%s a=%d%%", argb[2:], int(pct+0.5))
}
