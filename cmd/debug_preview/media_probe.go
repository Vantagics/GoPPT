package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"strings"
	"unicode/utf8"
)

// dumpMedia prints every media part in the package, so an image the renderer
// could not rasterise can be inspected for what it actually is.
//
// Text parts (SVG, and the XML some producers embed) are printed in full.
// Binary parts are summarised instead: dumping a PNG's bytes at a terminal
// produces a screenful of noise that hides the parts worth reading, and the
// useful facts about a raster image here are its size and format, not its
// payload.
//
//	go run ./cmd/debug_preview --media <file.pptx>
func dumpMedia(path string) {
	zr, err := zip.OpenReader(path)
	if err != nil {
		fmt.Printf("open: %v\n", err)
		return
	}
	defer zr.Close()
	for _, f := range zr.File {
		if !strings.Contains(f.Name, "/media/") || strings.HasSuffix(f.Name, "/") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			continue
		}
		data, _ := io.ReadAll(rc)
		rc.Close()
		if textLike(data) {
			fmt.Printf("=== %s (%d bytes, text) ===\n%s\n\n", f.Name, len(data), string(data))
			continue
		}
		fmt.Printf("=== %s (%d bytes, binary) ===\n  %s\n  first bytes: %s\n\n",
			f.Name, len(data), describeBinary(data), hexPreview(data, 16))
	}
}

// textLike reports whether data is plausibly text: no NUL byte, and valid
// UTF-8. Both tests are needed. A NUL is the classic binary tell, but a byte
// stream can also be NUL-free while being invalid UTF-8, and printing that
// garbles the report in a way that is easily mistaken for a decoding bug in the
// library rather than in the media part.
func textLike(data []byte) bool {
	if len(data) == 0 {
		return true
	}
	for _, b := range data {
		if b == 0 {
			return false
		}
	}
	return utf8.Valid(data)
}

// describeBinary names the format from its magic number, because "which format
// is this actually in" is the question a picture that failed to decode raises.
func describeBinary(data []byte) string {
	switch {
	case bytes.HasPrefix(data, []byte("\x89PNG\r\n\x1a\n")):
		return "PNG"
	case bytes.HasPrefix(data, []byte{0xFF, 0xD8, 0xFF}):
		return "JPEG"
	case bytes.HasPrefix(data, []byte("GIF87a")), bytes.HasPrefix(data, []byte("GIF89a")):
		return "GIF"
	case bytes.HasPrefix(data, []byte("BM")):
		return "BMP"
	case bytes.HasPrefix(data, []byte("II*\x00")), bytes.HasPrefix(data, []byte("MM\x00*")):
		return "TIFF"
	case bytes.HasPrefix(data, []byte("RIFF")):
		return "RIFF/WebP"
	case bytes.HasPrefix(data, []byte{0x01, 0x00, 0x00, 0x00}):
		return "EMF"
	case bytes.HasPrefix(data, []byte{0xD7, 0xCD, 0xC6, 0x9A}):
		return "WMF"
	default:
		return "unrecognised"
	}
}

func hexPreview(data []byte, n int) string {
	if len(data) < n {
		n = len(data)
	}
	var b strings.Builder
	for i := 0; i < n; i++ {
		if i > 0 {
			b.WriteByte(' ')
		}
		fmt.Fprintf(&b, "%02X", data[i])
	}
	return b.String()
}
