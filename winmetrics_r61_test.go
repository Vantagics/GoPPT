package gopresentation

import (
	"fmt"
	"testing"
)

// Round 61 helper: print the exact GDI win vertical metrics the renderer
// records for Calibri at the probe's sizes, so the float-layout phase can
// be fitted against the probe gold exports.
func TestZZWinMetricsPrint(t *testing.T) {
	if pickInstalledFont(NewFontCache()) == "" {
		t.Skip("no fonts installed")
	}
	fc := NewFontCache()
	for _, pt := range []float64{14, 16, 20, 28, 32} {
		px := pt * 12700.0 * (1600.0 / (720 * 12700.0)) // fontSizePixels at 1600px/720pt
		a, d, ok := fc.WinVerticalMetrics("Calibri", px, false, false)
		if !ok {
			t.Fatalf("no win metrics for Calibri at %fpx", px)
		}
		fmt.Printf("pt=%v px=%.6f ascF=%.6f descF=%.6f asc+desc=%.6f 1.2em=%.6f slackHalf=%.6f\n",
			pt, px, a, d, a+d, 1.2*px, (1.2*px-a-d)/2)
	}
}
