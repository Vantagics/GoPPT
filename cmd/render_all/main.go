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

	n := pres.GetSlideCount()
	fmt.Printf("Slide count: %d\n", n)
	for i := 0; i < n; i++ {
		outPath := filepath.Join(dst, fmt.Sprintf("slide%02d.png", i+1))
		fmt.Printf("Rendering slide %d -> %s\n", i+1, outPath)
		if err := pres.SaveSlideAsImage(i, outPath, opts); err != nil {
			fmt.Fprintf(os.Stderr, "slide %d: %v\n", i+1, err)
			continue
		}
		if info, err := os.Stat(outPath); err == nil {
			fmt.Printf("  OK: %d bytes\n", info.Size())
		} else {
			fmt.Printf("  MISSING: %v\n", err)
		}
	}

	fmt.Printf("Done. Rendered %d slides to %s\n", n, dst)
}
