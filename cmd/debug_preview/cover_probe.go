package main

import (
	"fmt"

	ppt "github.com/Vantagics/GoPPT"
)

// notdefProbeCandidates are codepoints unlikely to be mapped anywhere. The
// first one a font reports as absent gives that font's .notdef ink signature.
var notdefProbeCandidates = []rune{0xE000, 0xE001, 0xFDD0, 0x10FFFD, 0xF0000}

// probeCover answers the question the CJK font-selection code actually asks:
// does CoversRune say this font covers this rune, and does the rasterised ink
// agree?
//
// The pairs worth telling apart:
//
//	cov=N ink=104   absent     no glyph; the ink is the font's .notdef box
//	cov=Y ink=476   covered    a real glyph, whatever its size
//	cov=Y ink=0     blank      claims a glyph but draws nothing
//	cov=Y ink=<sig> SUSPECT    ink equals this font's .notdef signature, so the
//	                           font maps missing characters onto a real-looking
//	                           box glyph. That would defeat CoversRune, which
//	                           trusts "glyph index != 0" as proof of coverage.
//
// The SUSPECT row is the one that matters: every font-selection decision in the
// renderer rests on CoversRune being unable to be fooled this way.
//
//	go run ./cmd/debug_preview --cover U+2460 U+24B6 U+E000
func probeCover(args []string) {
	if len(args) == 0 {
		args = []string{"U+2460", "U+2461", "U+24B6", "U+24EA", "U+2776", "U+3220", "U+33A1"}
	}
	var runes []rune
	for _, raw := range args {
		r, ok := parseRuneArg(raw)
		if !ok {
			fmt.Printf("skipping %q: want a character or U+XXXX\n", raw)
			continue
		}
		runes = append(runes, r)
	}

	fc := ppt.NewFontCache()
	// The full production CJK fallback chain, plus a few Latin faces for
	// contrast. Mirroring the chain matters: a symbol may be drawable by a font
	// the renderer will never consider, which changes nothing, or by one it
	// will, which changes the answer.
	names := []string{
		"Microsoft YaHei", "SimSun", "SimHei", "NSimSun",
		"Yu Gothic", "Meiryo", "MS Gothic",
		"Malgun Gothic", "Gulim",
		"Noto Sans CJK SC", "Noto Sans SC", "WenQuanYi Micro Hei",
		"Arial", "Calibri", "Segoe UI", "DengXian",
		// The symbol/emoji chain, in production order.
		"Segoe UI Emoji", "Apple Color Emoji", "Noto Color Emoji",
		"Segoe UI Symbol", "Noto Sans Symbols 2", "Noto Sans Symbols",
		"DejaVu Sans",
	}

	// Establish each font's .notdef signature before judging any of the samples.
	notdef := map[string]int{}
	for _, name := range names {
		face := fc.GetFace(name, 48, false, false)
		if face == nil {
			continue
		}
		for _, r := range notdefProbeCandidates {
			if fc.CoversRune(name, false, false, r) {
				continue
			}
			if ink := inkCount(face, string(r)); ink > 0 {
				notdef[name] = ink
				break
			}
		}
	}

	for _, r := range runes {
		fmt.Printf("\n%s\n", runeLabel(r))
		for _, name := range names {
			face := fc.GetFace(name, 48, false, false)
			if face == nil {
				fmt.Printf("  %-18s (no face)\n", name)
				continue
			}
			cov := fc.CoversRune(name, false, false, r)
			ink := inkCount(face, string(r))
			verdict := coverVerdict(cov, ink, notdef[name])
			fmt.Printf("  %-18s cov=%-5v ink=%-5d %s\n", name, yesNo(cov), ink, verdict)
		}
	}
}

func yesNo(b bool) string {
	if b {
		return "Y"
	}
	return "N"
}

// coverVerdict classifies a (CoversRune, ink) pair. sig is the font's .notdef
// ink signature, or 0 when it could not be established.
func coverVerdict(cov bool, ink, sig int) string {
	switch {
	case cov && ink == 0:
		return "blank  (claims a glyph, draws nothing)"
	case cov && sig > 0 && ink == sig:
		return fmt.Sprintf("SUSPECT (ink matches .notdef signature %d)", sig)
	case cov:
		return "covered"
	case ink > 0:
		return "absent (.notdef box)"
	default:
		return "absent (nothing drawn)"
	}
}
