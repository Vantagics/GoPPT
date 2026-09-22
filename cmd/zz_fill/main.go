package main

import (
	"fmt"
	"os"

	gopresentation "github.com/Vantagics/GoPPT"
)

func main() {
	path := os.Args[1]
	slideIdx := 0
	fmt.Sscanf(os.Args[2], "%d", &slideIdx)

	pres, err := gopresentation.Open(path)
	if err != nil {
		fmt.Println("open:", err)
		return
	}
	slides := pres.Slides()
	if slideIdx < 1 || slideIdx > len(slides) {
		fmt.Println("bad slide index")
		return
	}
	sl := slides[slideIdx-1]
	for i, sh := range sl.GetShapes() {
		rt, ok := sh.(interface{ GetParagraphs() []*gopresentation.Paragraph })
		if !ok {
			continue
		}
		_ = rt
		name := "?"
		if n, ok := sh.(interface{ GetName() string }); ok {
			name = n.GetName()
		}
		var fillDesc string
		if f, ok := sh.(interface{ GetFill() *gopresentation.Fill }); ok {
			if fl := f.GetFill(); fl != nil {
				switch fl.Type {
				case gopresentation.FillSolid:
					fillDesc = fmt.Sprintf("solid %s", fl.Color.ARGB)
				case gopresentation.FillGradientLinear:
					fillDesc = fmt.Sprintf("grad %s -> %s rot=%d mid=%s@%d", fl.Color.ARGB, fl.EndColor.ARGB, fl.Rotation, fl.MidColor.ARGB, fl.MidPos)
				case gopresentation.FillNone:
					fillDesc = "none"
				default:
					fillDesc = fmt.Sprintf("type=%d", fl.Type)
				}
			}
		}
		fmt.Printf("[%d] %s fill=%s\n", i, name, fillDesc)
		if bd, ok := sh.(interface{ GetBorder() *gopresentation.Border }); ok {
			if b := bd.GetBorder(); b != nil {
				fmt.Printf("      border=%s w=%dpt style=%s\n", b.Color.ARGB, b.Width, b.Style)
			}
		}
	}
}
