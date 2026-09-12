package main

import "testing"

// TestTextLikeKeepsBinaryOutOfTheReport guards the two independent ways a part
// can be binary.
//
// Either test alone is insufficient, and the failure is not cosmetic: a raster
// part printed as text floods the report with noise that hides the parts worth
// reading.
func TestTextLikeKeepsBinaryOutOfTheReport(t *testing.T) {
	tests := []struct {
		name string
		data []byte
		want bool
	}{
		{"empty part", nil, true},
		{"svg", []byte(`<svg viewBox="0 0 16 16"><path d="M0 0h16v16H0z"/></svg>`), true},
		{"xml", []byte("<?xml version=\"1.0\"?><a/>"), true},
		{"CJK text", []byte("组织的驱动系统"), true},
		{"embedded NUL", []byte{'a', 0x00, 'b'}, false},
		// A PNG header has no NUL, so only the UTF-8 test rejects it. Without
		// that test this input is classified as text and dumped verbatim.
		{"PNG header, no NUL", []byte("\x89PNG\r\n\x1a\n"), false},
		// Truncated two-byte sequence: also NUL-free, also invalid UTF-8.
		{"invalid UTF-8", []byte{0xC3, 0x28}, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := textLike(tt.data); got != tt.want {
				t.Errorf("textLike(%q) = %v, want %v", tt.data, got, tt.want)
			}
		})
	}
}

func TestDescribeBinaryNamesTheFormat(t *testing.T) {
	tests := []struct {
		name string
		data []byte
		want string
	}{
		{"png", []byte("\x89PNG\r\n\x1a\n"), "PNG"},
		{"jpeg", []byte{0xFF, 0xD8, 0xFF, 0xE0}, "JPEG"},
		{"gif", []byte("GIF89a"), "GIF"},
		{"bmp", []byte("BM"), "BMP"},
		{"tiff little endian", []byte("II*\x00"), "TIFF"},
		{"tiff big endian", []byte("MM\x00*"), "TIFF"},
		{"webp", []byte("RIFF"), "RIFF/WebP"},
		{"emf", []byte{0x01, 0x00, 0x00, 0x00}, "EMF"},
		{"wmf", []byte{0xD7, 0xCD, 0xC6, 0x9A}, "WMF"},
		{"svg is not a raster format", []byte("<svg/>"), "unrecognised"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := describeBinary(tt.data); got != tt.want {
				t.Errorf("describeBinary(%x) = %q, want %q", tt.data, got, tt.want)
			}
		})
	}
}

func TestHexPreviewTruncates(t *testing.T) {
	if got := hexPreview([]byte{0x01, 0xAB, 0xFF}, 16); got != "01 AB FF" {
		t.Errorf("short input = %q, want %q", got, "01 AB FF")
	}
	if got := hexPreview([]byte{0x01, 0x02, 0x03, 0x04}, 2); got != "01 02" {
		t.Errorf("truncated = %q, want %q", got, "01 02")
	}
	if got := hexPreview(nil, 16); got != "" {
		t.Errorf("empty = %q, want %q", got, "")
	}
}
