package main

import (
	"fmt"
	"os"

	gopresentation "github.com/Vantagics/GoPPT"
)

func main() {
	reader, _ := gopresentation.NewReader(gopresentation.ReaderPowerPoint2007)
	pres, err := reader.Read("test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Read: %v\n", err)
		os.Exit(1)
	}

	opts := gopresentation.DefaultRenderOptions()
	opts.Width = 1920
	outPath := "cmd/render_slide7/slide7.png"
	if err := pres.SaveSlideAsImage(6, outPath, opts); err != nil {
		fmt.Fprintf(os.Stderr, "render: %v\n", err)
		os.Exit(1)
	}
	fmt.Println("Rendered:", outPath)
}
