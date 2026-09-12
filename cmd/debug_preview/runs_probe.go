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
				rt, ok := s.(*ppt.RichTextShape)
				if !ok {
					continue
				}
				fmt.Printf("slide%02d %s box=(%d,%d %dx%d)\n", i+1,
					strings.Repeat("  ", depth), rt.GetOffsetX(), rt.GetOffsetY(), rt.GetWidth(), rt.GetHeight())
				for pi, p := range rt.GetParagraphs() {
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
						fmt.Printf("  p%d run %s sz=%d bold=%v latin=%q ea=%q color=%s\n",
							pi, quoteASCII(tr.GetText()), f.Size, f.Bold, f.Name, f.NameEA, describeARGB(f.Color.ARGB))
					}
				}
			}
		}
		walk(slide.GetShapes(), 0)
	}
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
