package main

import (
	"fmt"
	"image"
	"image/color"
	"image/draw"
	"strings"

	ppt "github.com/Vantagics/GoPPT"
	"golang.org/x/image/font"
	"golang.org/x/image/math/fixed"
)

// inkCount rasterises s with the given face and counts non-background pixels.
// A real glyph has ink; a missing-glyph box (.notdef) has far less, and a face
// with no glyph at all draws nothing.
func inkCount(face font.Face, s string) int {
	const w, h = 160, 80
	img := image.NewRGBA(image.Rect(0, 0, w, h))
	draw.Draw(img, img.Bounds(), image.NewUniform(color.RGBA{255, 255, 255, 255}), image.Point{}, draw.Src)
	d := font.Drawer{
		Dst:  img,
		Src:  image.NewUniform(color.RGBA{0, 0, 0, 255}),
		Face: face,
		Dot:  fixed.P(5, 55),
	}
	d.DrawString(s)
	n := 0
	for y := 0; y < h; y++ {
		for x := 0; x < w; x++ {
			if img.Pix[y*img.Stride+x*4] < 200 {
				n++
			}
		}
	}
	return n
}

// probeInk checks whether the faces the renderer would pick can actually draw
// CJK. Pass sample characters as arguments to test other scripts.
//
//	A real glyph has a different ink count for each distinct character. A face
//	missing the glyph draws its .notdef box, so every sample gets the SAME
//	count. That uniformity is the tell for "no glyph here".
//
// Samples may be written literally or as ASCII "U+XXXX" escapes. Prefer the
// escapes when driving this from a shell whose codepage may mangle literals.
//
//	go run ./cmd/debug_preview --ink                    # the default samples
//	go run ./cmd/debug_preview --ink U+2460 U+2461 组
func probeInk(samples []string) {
	if len(samples) == 0 {
		samples = []string{"组", "织", "U+2460", "U+2461", "U+2462", "U+2463"}
	}
	type sample struct {
		label string
		text  string
	}
	parsed := make([]sample, 0, len(samples))
	for _, raw := range samples {
		ch, ok := parseRuneArg(raw)
		if !ok {
			fmt.Printf("skipping %q: want a character or U+XXXX\n", raw)
			continue
		}
		parsed = append(parsed, sample{label: runeLabel(ch), text: string(ch)})
	}
	fc := ppt.NewFontCache()
	names := []string{"Microsoft YaHei", "SimSun", "Arial", "Calibri", "Segoe UI", "DengXian"}
	for _, name := range names {
		for _, style := range []struct {
			label        string
			bold, italic bool
		}{{"regular", false, false}, {"bold", true, false}} {
			face := fc.GetFace(name, 48, style.bold, style.italic)
			if face == nil {
				fmt.Printf("%-18s %-8s MISS\n", name, style.label)
				continue
			}
			var parts []string
			for _, s := range parsed {
				parts = append(parts, fmt.Sprintf("%s=%-5d", s.label, inkCount(face, s.text)))
			}
			fmt.Printf("%-18s %-8s %s\n", name, style.label, strings.Join(parts, " "))
		}
	}
}

// parseRuneArg accepts a single character, or the ASCII spellings "U+2460",
// "u2460" and "2460" for a code point.
func parseRuneArg(s string) (rune, bool) {
	rs := []rune(s)
	if len(rs) == 1 && rs[0] != 'U' && rs[0] != 'u' {
		return rs[0], true
	}
	hex := strings.TrimPrefix(strings.TrimPrefix(s, "U+"), "u+")
	hex = strings.TrimPrefix(strings.TrimPrefix(hex, "U"), "u")
	if hex == "" || len(hex) > 8 {
		return 0, false
	}
	// Only a pure hex string is a code point; anything else is a mistake.
	v := 0
	for _, c := range hex {
		var d int
		switch {
		case c >= '0' && c <= '9':
			d = int(c - '0')
		case c >= 'a' && c <= 'f':
			d = int(c-'a') + 10
		case c >= 'A' && c <= 'F':
			d = int(c-'A') + 10
		default:
			return 0, false
		}
		v = v*16 + d
	}
	if v == 0 || v > 0x10FFFF {
		return 0, false
	}
	return rune(v), true
}

// runeLabel renders a rune as pure ASCII so a console codepage cannot corrupt
// the report. Printable ASCII is shown as-is, everything else as U+XXXX.
func runeLabel(r rune) string {
	if r >= 0x21 && r <= 0x7E {
		return string(r)
	}
	return fmt.Sprintf("U+%04X", r)
}
