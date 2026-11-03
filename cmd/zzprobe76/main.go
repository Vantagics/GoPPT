// Command zzprobe76 prints per-slide backgrounds and, for slide 7, the shape
// list — a multi-master theme round's view of which rung each slide reached.
//
//	zzprobe76 <deck.pptx>
package main

import (
	"fmt"
	"os"

	gopresentation "github.com/Vantagics/GoPPT"
)

func main() {
	pres, err := gopresentation.Open(os.Args[1])
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	for i, s := range pres.GetAllSlides() {
		bg := s.GetBackground()
		desc := "nil"
		if bg != nil {
			desc = fmt.Sprintf("type=%v color=%v", bg.Type, bg.Color)
		}
		fmt.Printf("slide%02d shapes=%d bg=%s\n", i+1, len(s.GetShapes()), desc)
		if i != 6 {
			continue
		}
		for j, sh := range s.GetShapes() {
			name := ""
			if v, ok := sh.(interface{ GetName() string }); ok {
				name = v.GetName()
			}
			geo := ""
			if v, ok := sh.(interface {
				GetOffsetX() int
				GetOffsetY() int
			}); ok {
				geo = fmt.Sprintf(" off=(%d,%d)", v.GetOffsetX(), v.GetOffsetY())
			}
			fmt.Printf("   [%d] %T %s%s\n", j, sh, name, geo)
		}
	}
}
