package main

import (
	"fmt"
	"os"
	"path/filepath"

	gopresentation "github.com/VantageDataChat/GoPPT"
)

func main() {
	src := "test.pptx"
	dst := `C:\Users\ma139\AppData\Local\Temp\goppt_render_test`

	if err := os.MkdirAll(dst, 0750); err != nil {
		fmt.Fprintf(os.Stderr, "mkdir: %v\n", err)
		os.Exit(1)
	}

	reader, err := gopresentation.NewReader(gopresentation.ReaderPowerPoint2007)
	if err != nil {
		fmt.Fprintf(os.Stderr, "new reader: %v\n", err)
		os.Exit(1)
	}

	pres, err := reader.Read(src)
	if err != nil {
		fmt.Fprintf(os.Stderr, "read: %v\n", err)
		os.Exit(1)
	}

	opts := gopresentation.DefaultRenderOptions()
	opts.Width = 1920

	pattern := filepath.Join(dst, "slide%02d.png")
	if err := pres.SaveSlidesAsImages(pattern, opts); err != nil {
		fmt.Fprintf(os.Stderr, "render: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("Rendered %d slides to %s\n", pres.GetSlideCount(), dst)
}
