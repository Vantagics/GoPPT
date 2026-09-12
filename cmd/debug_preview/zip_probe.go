package main

import (
	"archive/zip"
	"fmt"
	"strings"

	ppt "github.com/Vantagics/GoPPT"
)

// listZip prints the entries of a .pptx and, side by side, what the reader
// makes of it — a slide part the reader never counted is the quickest way to
// see that a slide was dropped.
//
//	go run ./cmd/debug_preview --ls <file.pptx>
func listZip(path string) {
	zr, err := zip.OpenReader(path)
	if err != nil {
		fmt.Printf("open: %v\n", err)
		return
	}
	defer zr.Close()

	slides := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			slides++
		}
	}
	fmt.Printf("zip entries      : %d\n", len(zr.File))
	fmt.Printf("slide parts      : %d\n", slides)

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
	fmt.Printf("reader slideCount: %d\n", pres.GetSlideCount())
	fmt.Println()
	for _, f := range zr.File {
		fmt.Printf("  %-46s %7d\n", f.Name, f.UncompressedSize64)
	}
}
