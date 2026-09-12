package gopresentation

import "testing"

// WMF rasterisation — the canvas must be bounded by the file, not by the file's
// claims.
//
// decodeWMFDIB sizes its canvas from META_SETWINDOWEXT, which is two signed
// 16-bit values read straight out of the metafile, and multiplied the result by
// a fixed 4x quality factor. A file claiming a 30000x30000 logical extent
// therefore asked image.NewRGBA for a 12000x12000 buffer — 576 MB, and 68 GB at
// the maximum the two int16 fields allow. image.NewRGBA does not return an error
// for that, it dies in the allocator, and the recover boundary safety.go
// installs cannot catch an allocation failure. Every other dimension in this
// path is already checked: parseDIB refuses anything over 4096 and the EMF vector
// path clamps its canvas to 2000 px. Only the WMF canvas had no ceiling.

// wmfWithWindowExt builds the smallest metafile that reaches the render-sizing
// code: an 18-byte header, a META_SETWINDOWEXT record giving the logical extent,
// and one ExtTextOut record carrying "AB". The text record is what makes
// decodeWMFDIB carry on rather than return early with nothing to draw.
func wmfWithWindowExt(winW, winH int) []byte {
	u16 := func(v int) []byte { return []byte{byte(v), byte(v >> 8)} }
	// A WMF record is a 4-byte size in words, a 2-byte function and a body.
	rec := func(function int, body []byte) []byte {
		words := (6 + len(body)) / 2
		out := []byte{byte(words), byte(words >> 8), 0, 0}
		out = append(out, u16(function)...)
		return append(out, body...)
	}

	// META_SETWINDOWEXT is followed by the y extent and then the x extent.
	setWindowExt := rec(0x020C, append(u16(winH), u16(winW)...))
	// ExtTextOut: y, x, count, options, then the string (options 0 means no
	// rectangle precedes it).
	body := append(u16(0), u16(0)...) // y, x
	body = append(body, u16(2)...)    // character count
	body = append(body, u16(0)...)    // options
	body = append(body, 'A', 'B')     // the string itself
	extTextOut := rec(0x0A32, body)

	data := []byte{0x01, 0x00, 0x09, 0x00}
	data = append(data, make([]byte, 14)...) // rest of the WMF header
	data = append(data, setWindowExt...)
	return append(data, extTextOut...)
}

// TestWMFCanvasIsBoundedByTheWindowExtent is the regression: a metafile asking
// for a 700x700 logical extent must not produce a 2800x2800 canvas.
func TestWMFCanvasIsBoundedByTheWindowExtent(t *testing.T) {
	const maxMetafileDim = 2000

	img := decodeMetafileBitmap(wmfWithWindowExt(700, 700), NewFontCache())
	if img == nil {
		t.Fatal("a metafile with a valid ExtTextOut record rendered nothing")
	}
	bounds := img.Bounds()
	if bounds.Dx() > maxMetafileDim || bounds.Dy() > maxMetafileDim {
		t.Errorf("a 700x700 logical metafile produced a %dx%d canvas, over the %d px ceiling; the window extent is untrusted input",
			bounds.Dx(), bounds.Dy(), maxMetafileDim)
	}
}

// TestWMFKeepsItsUpscaleForOrdinarySizes is the control: the ceiling must not
// cost small metafiles their 4x quality multiplier.
func TestWMFKeepsItsUpscaleForOrdinarySizes(t *testing.T) {
	img := decodeMetafileBitmap(wmfWithWindowExt(100, 80), NewFontCache())
	if img == nil {
		t.Fatal("a metafile with a valid ExtTextOut record rendered nothing")
	}
	bounds := img.Bounds()
	if bounds.Dx() != 400 || bounds.Dy() != 320 {
		t.Errorf("a 100x80 logical metafile produced %dx%d, want 400x320 (the 4x upscale)", bounds.Dx(), bounds.Dy())
	}
}
