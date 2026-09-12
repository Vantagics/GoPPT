// Command debug_preview inspects a .pptx, and the preview images rendered from
// it. It is the harness the renderer is accepted against: render every slide,
// then ask the model, the package XML and the pixels what actually happened, so
// a fidelity bug can be attributed rather than guessed at.
//
// Primary mode — render a deck and print a fidelity report:
//
//	go run ./cmd/debug_preview <file.pptx> <outdir> [width]
//
// Inspection modes:
//
//	--shapes <file.pptx>              per-slide shape inventory, picture sizes
//	--runs <file.pptx> [slide]        text runs with resolved font and colour
//	--ls <file.pptx>                  package entries, next to the reader's view
//	--part <file.pptx> <name> [sub]   one part, or the hits around a substring
//	--pretty <file.pptx> <name>       one part, one element per line
//	--media <file.pptx>               every media part, text ones in full
//	--pix <file.png> [x y w h]        colour histogram and ASCII map of a region
//	--crop <in.png> <out.png> [x y w h] [zoom]
//	                                  zoom a region so small text can be read
//	--gallery <dir> <out.html> [width] [title]
//	                                  self-contained HTML sheet of a render
//	--fonts                           which known font names resolve here
//	--ink [char|U+XXXX ...]           ink per sample per face; equal counts
//	                                  across different characters mean .notdef
//	--cover [char|U+XXXX ...]         cross-check CoversRune against the ink
//
// Every mode that reports text prints ASCII only (\uXXXX escapes), because the
// host console decodes output with the system codepage and would otherwise
// garble the very code points being investigated.
package main

import (
	"fmt"
	"image"
	"image/png"
	"io"
	"os"
	"path/filepath"
	"strconv"

	ppt "github.com/Vantagics/GoPPT"
)

// usage writes the mode list. It goes to stdout for an explicit request and to
// stderr when it is a reaction to a malformed command line.
func usage(w io.Writer) {
	fmt.Fprint(w, `usage: debug_preview <file.pptx> <outdir> [width]
       debug_preview --shapes <file.pptx>
       debug_preview --runs <file.pptx> [slide-number]
       debug_preview --ls <file.pptx>
       debug_preview --part <file.pptx> <part-name> [substring]
       debug_preview --pretty <file.pptx> <part-name>
       debug_preview --media <file.pptx>
       debug_preview --pix <file.png> [x y w h]
       debug_preview --crop <in.png> <out.png> [x y w h] [zoom]
       debug_preview --gallery <dir> <out.html> [thumb-width] [title]
       debug_preview --fonts
       debug_preview --ink [char|U+XXXX ...]
       debug_preview --cover [char|U+XXXX ...]
`)
}

func main() {
	if len(os.Args) >= 2 {
		switch os.Args[1] {
		case "-h", "--help", "help":
			usage(os.Stdout)
			return
		}
	}
	if len(os.Args) >= 2 && os.Args[1] == "--fonts" {
		probeFonts()
		return
	}
	if len(os.Args) >= 2 && os.Args[1] == "--ink" {
		probeInk(os.Args[2:])
		return
	}
	if len(os.Args) >= 2 && os.Args[1] == "--cover" {
		probeCover(os.Args[2:])
		return
	}
	if len(os.Args) >= 3 && os.Args[1] == "--gallery" {
		writeGallery(os.Args[2:])
		return
	}
	if len(os.Args) >= 3 && os.Args[1] == "--ls" {
		listZip(os.Args[2])
		return
	}
	if len(os.Args) >= 3 && os.Args[1] == "--shapes" {
		dumpShapes(os.Args[2])
		return
	}
	if len(os.Args) >= 3 && os.Args[1] == "--media" {
		dumpMedia(os.Args[2])
		return
	}
	if len(os.Args) >= 4 && os.Args[1] == "--part" {
		only := ""
		if len(os.Args) >= 5 {
			only = os.Args[4]
		}
		dumpPart(os.Args[2], os.Args[3], only)
		return
	}
	if len(os.Args) >= 4 && os.Args[1] == "--pretty" {
		dumpPartPretty(os.Args[2], os.Args[3])
		return
	}
	if len(os.Args) >= 3 && os.Args[1] == "--pix" {
		inspectPixels(os.Args[2:])
		return
	}
	if len(os.Args) >= 3 && os.Args[1] == "--crop" {
		cropAndZoom(os.Args[2:])
		return
	}
	if len(os.Args) >= 3 && os.Args[1] == "--runs" {
		only := 0
		if len(os.Args) >= 4 {
			only, _ = strconv.Atoi(os.Args[3])
		}
		probeRuns(os.Args[2], only)
		return
	}
	if len(os.Args) < 3 {
		usage(os.Stderr)
		os.Exit(2)
	}
	path := os.Args[1]
	outDir := os.Args[2]
	width := 1600
	if len(os.Args) > 3 {
		if v, err := strconv.Atoi(os.Args[3]); err == nil && v > 0 {
			width = v
		}
	}
	if err := os.MkdirAll(outDir, 0o755); err != nil {
		fatal("mkdir %s: %v", outDir, err)
	}

	reader, err := ppt.NewReader(ppt.ReaderPowerPoint2007)
	if err != nil {
		fatal("new reader: %v", err)
	}
	pres, err := reader.Read(path)
	if err != nil {
		fatal("read %s: %v", path, err)
	}

	layout := pres.GetLayout()
	fmt.Printf("file      : %s\n", path)
	fmt.Printf("slides    : %d\n", pres.GetSlideCount())
	fmt.Printf("layout    : %d x %d EMU  (%.4f x %.4f in, ratio %.4f)\n",
		layout.CX, layout.CY, float64(layout.CX)/914400, float64(layout.CY)/914400,
		float64(layout.CX)/float64(layout.CY))

	// Constructs the reader could not represent. A non-empty list means part of
	// the deck is drawn as an amber placeholder rather than as real content.
	if bad := pres.UnsupportedShapes(); len(bad) > 0 {
		fmt.Printf("unsupported: %d\n", len(bad))
		for i, s := range bad {
			fmt.Printf("  [%d] %s  ct=%q  box=(%d,%d %dx%d)\n", i, s.Label(), s.GetContentType(),
				s.GetOffsetX(), s.GetOffsetY(), s.GetWidth(), s.GetHeight())
		}
	} else {
		fmt.Printf("unsupported: none\n")
	}

	diag := ppt.NewFontDiagnostics()
	opts := ppt.DefaultRenderOptions()
	opts.Width = width
	opts.FontDiagnostics = diag
	// The preview should be faithful, so no Draft.

	imgs, err := pres.SlidesToImages(opts)
	if err != nil {
		fatal("render: %v", err)
	}

	fmt.Printf("rendered  : %d images at %d px wide\n", len(imgs), width)
	for i, img := range imgs {
		b := img.Bounds()
		name := fmt.Sprintf("slide%02d.png", i+1)
		if err := writePNG(filepath.Join(outDir, name), img); err != nil {
			fmt.Fprintf(os.Stderr, "save %s: %v\n", name, err)
			continue
		}
		fmt.Printf("  %-14s %dx%d\n", name, b.Dx(), b.Dy())
	}

	// A picture whose data never made it out of the package is drawn as an
	// amber placeholder; say so here rather than making the reader spot it in
	// the PNG. Reporting it from the model means there are no false positives
	// from deck colours that happen to match the placeholder's.
	reportEmptyPictures(pres)

	if summary := diag.Summary(); summary != "" {
		fmt.Printf("fonts     : %s\n", summary)
		for _, u := range diag.Usages() {
			fmt.Printf("  %-40s %s\n", u.Requested, u.Kind)
		}
	} else {
		fmt.Printf("fonts     : all requests satisfied exactly\n")
	}
}

// reportEmptyPictures lists every picture whose image data is empty, which the
// renderer draws as a labelled placeholder.
//
// Pictures nested in groups are reported too: a group is where a picture is
// most likely to be overlooked, and the placeholder is inside the group's box.
func reportEmptyPictures(pres *ppt.Presentation) {
	var empty []string
	for i, slide := range pres.GetAllSlides() {
		var walk func(shapes []ppt.Shape)
		walk = func(shapes []ppt.Shape) {
			for _, s := range shapes {
				switch v := s.(type) {
				case *ppt.DrawingShape:
					if len(v.GetImageData()) == 0 {
						empty = append(empty, fmt.Sprintf(
							"slide%02d box=(%d,%d %dx%d) mime=%q",
							i+1, v.GetOffsetX(), v.GetOffsetY(), v.GetWidth(), v.GetHeight(), v.GetMimeType()))
					}
				case *ppt.GroupShape:
					walk(v.GetShapes())
				}
			}
		}
		walk(slide.GetShapes())
	}
	if len(empty) == 0 {
		fmt.Printf("pictures  : every picture has image data\n")
		return
	}
	fmt.Printf("pictures  : %d drawn as placeholder (no image data)\n", len(empty))
	for _, e := range empty {
		fmt.Printf("  %s\n", e)
	}
}

func writePNG(path string, img image.Image) error {
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()
	return png.Encode(f, img)
}

func fatal(format string, args ...any) {
	fmt.Fprintf(os.Stderr, "debug_preview: "+format+"\n", args...)
	os.Exit(1)
}
