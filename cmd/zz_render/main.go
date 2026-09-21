// Command zz_render renders every slide of a deck to PNG files named
// slideNN.png, the naming the PowerPoint COM export (cmd/zz_cmp/
// export_office.ps1) and the comparison tool cmd/zz_cmp pair on.
//
//	zz_render <deck.pptx> <outdir> [width]
//
// It exists so a fidelity round can re-render the comparison deck without
// touching the hardcoded cmd/render_all.
package main

import (
	"fmt"
	"os"
	"path/filepath"

	gopresentation "github.com/Vantagics/GoPPT"
)

func main() {
	if len(os.Args) < 3 {
		fmt.Fprintln(os.Stderr, "usage: zz_render <deck.pptx> <outdir> [width]")
		os.Exit(2)
	}
	deck, outDir := os.Args[1], os.Args[2]
	width := 1600
	if len(os.Args) > 3 {
		if _, err := fmt.Sscanf(os.Args[3], "%d", &width); err != nil {
			fmt.Fprintln(os.Stderr, "bad width:", os.Args[3])
			os.Exit(2)
		}
	}

	reader, err := gopresentation.NewReader(gopresentation.ReaderPowerPoint2007)
	if err != nil {
		fmt.Fprintf(os.Stderr, "new reader: %v\n", err)
		os.Exit(1)
	}
	pres, err := reader.Read(deck)
	if err != nil {
		fmt.Fprintf(os.Stderr, "read %s: %v\n", deck, err)
		os.Exit(1)
	}
	if err := os.MkdirAll(outDir, 0o755); err != nil {
		fmt.Fprintf(os.Stderr, "mkdir %s: %v\n", outDir, err)
		os.Exit(1)
	}

	opts := gopresentation.DefaultRenderOptions()
	opts.Width = width
	n := pres.GetSlideCount()
	rendered := 0
	for i := 0; i < n; i++ {
		out := filepath.Join(outDir, fmt.Sprintf("slide%02d.png", i+1))
		if err := pres.SaveSlideAsImage(i, out, opts); err != nil {
			fmt.Fprintf(os.Stderr, "slide %d: %v\n", i+1, err)
			continue
		}
		rendered++
	}
	fmt.Printf("rendered %d/%d slides at width %d to %s\n", rendered, n, width, outDir)
}
