package main

import (
	"fmt"

	ppt "github.com/Vantagics/GoPPT"
)

// probeFonts prints which of the well-known font names the FontCache can
// resolve on this machine. It answers "is CJK tofu a missing font, or a
// resolution bug?" without guessing.
//
//	go run ./cmd/debug_preview --fonts
func probeFonts() {
	fc := ppt.NewFontCache()
	names := []string{
		"Microsoft YaHei", "SimSun", "SimHei", "NSimSun",
		"Yu Gothic", "Meiryo", "MS Gothic",
		"Noto Sans CJK SC", "Noto Sans SC", "WenQuanYi Micro Hei",
		"Arial", "Helvetica", "DejaVu Sans",
		"Segoe UI", "Calibri", "Inter", "Poppins", "Roboto", "Lato",
		"Source Han Sans SC", "Noto Sans", "HarmonyOS Sans SC",
		"Alibaba PuHuiTi", "PingFang SC", "Microsoft JhengHei",
		"DengXian", "FangSong", "KaiTi", "Cascadia Code", "Consolas",
	}
	state := func(x any) string {
		if x == nil {
			return "MISS"
		}
		return "ok"
	}
	fmt.Printf("%-24s %-6s %-6s\n", "requested", "face", "meas")
	for _, n := range names {
		fmt.Printf("%-24s %-6s %-6s\n", n,
			state(fc.GetFace(n, 24, false, false)),
			state(fc.GetMeasureFace(n, 24, false, false)))
	}
}
