package main

import (
	"fmt"
	"os"

	gopresentation "github.com/VantageDataChat/GoPPT"
)

func main() {
	reader, err := gopresentation.NewReader(gopresentation.ReaderPowerPoint2007)
	if err != nil {
		fmt.Fprintf(os.Stderr, "new reader: %v\n", err)
		os.Exit(1)
	}
	pres, err := reader.Read("test2.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "read: %v\n", err)
		os.Exit(1)
	}
	opts := gopresentation.DefaultRenderOptions()
	opts.Width = 1920
	n := pres.GetSlideCount()
	fmt.Printf("Total slides: %d\n", n)
	for i := 0; i < n; i++ {
		outPath := fmt.Sprintf("cmd/render_all/slide%02d.png", i+1)
		if err := pres.SaveSlideAsImage(i, outPath, opts); err != nil {
			fmt.Fprintf(os.Stderr, "slide %d: %v\n", i+1, err)
			continue
		}
		fmt.Printf("OK: %s\n", outPath)
	}
}
