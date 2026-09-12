package main

import (
	"fmt"

	ppt "github.com/Vantagics/GoPPT"
)

// dumpShapes prints a per-slide inventory of shape types and geometry, so a
// preview can be cross-checked against what the document actually contains
// (e.g. "slide 12 has a table, but nothing table-like was drawn").
//
//	go run ./cmd/debug_preview --shapes <file.pptx>
func dumpShapes(path string) {
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
	totals := map[string]int{}
	for i, s := range pres.GetAllSlides() {
		if s == nil {
			continue
		}
		counts := map[string]int{}
		var walk func(shapes []ppt.Shape, depth int)
		walk = func(shapes []ppt.Shape, depth int) {
			for _, sh := range shapes {
				name := sh.GetType().String()
				counts[name]++
				totals[name]++
				if name == "Unsupported" {
					if u, ok := sh.(*ppt.UnsupportedShape); ok {
						name += "(" + u.GetReason() + ")"
					}
				}
				if depth < 2 {
					extra := ""
					if pic, ok := sh.(*ppt.DrawingShape); ok {
						extra = fmt.Sprintf(" mime=%q bytes=%d path=%q",
							pic.GetMimeType(), len(pic.GetImageData()), pic.GetPath())
					}
					fmt.Printf("  slide%02d %s%s box=(%d,%d %dx%d)%s\n", i+1,
						indent(depth), name,
						sh.GetOffsetX(), sh.GetOffsetY(), sh.GetWidth(), sh.GetHeight(), extra)
				}
				if g, ok := sh.(*ppt.GroupShape); ok {
					walk(g.GetShapes(), depth+1)
				}
			}
		}
		walk(s.GetShapes(), 0)
		fmt.Printf("slide%02d: %v\n", i+1, counts)
	}
	fmt.Printf("totals: %v\n", totals)
}

func indent(n int) string {
	s := ""
	for i := 0; i < n; i++ {
		s += "  "
	}
	return s
}
