package main

// Prints every shape of the requested slides: kind, name, position, and each
// paragraph's runs with their field types. Diagnostic for the slidenum work.

import (
	"fmt"
	"os"
	"strconv"
	"strings"

	gopresentation "github.com/Vantagics/GoPPT"
)

func main() {
	deck := os.Args[1]
	r, err := gopresentation.NewReader(gopresentation.ReaderPowerPoint2007)
	if err != nil {
		panic(err)
	}
	pres, err := r.Read(deck)
	if err != nil {
		panic(err)
	}
	for _, sn := range os.Args[2:] {
		idx, _ := strconv.Atoi(sn)
		slide := pres.GetAllSlides()[idx-1]
		fmt.Printf("=== slide %d ===\n", idx)
		for i, sh := range slide.GetShapes() {
			ox, oy, w, h := sh.GetOffsetX(), sh.GetOffsetY(), sh.GetWidth(), sh.GetHeight()
			fmt.Printf(" [%d] %T name=%q off=(%d,%d) ext=(%d,%d)\n", i, sh, sh.GetName(), ox, oy, w, h)
			if rt, ok := sh.(interface {
				GetParagraphs() []*gopresentation.Paragraph
			}); ok {
				for pi, p := range rt.GetParagraphs() {
					var parts []string
					for _, el := range p.GetElements() {
						if tr, ok := el.(*gopresentation.TextRun); ok {
							parts = append(parts, strconv.Quote(tr.GetText())+"(fld="+strconv.Quote(tr.GetFieldType())+")")
						}
					}
					if len(parts) > 0 {
						fmt.Printf("     p%d: %s\n", pi, strings.Join(parts, " "))
					}
				}
			}
		}
	}
}
