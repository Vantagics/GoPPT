package gopresentation

import (
	"testing"
)

func TestProbeTexResolve(t *testing.T) {
	pres, err := Open("F:/pptx_corpus/web/themes.pptx")
	if err != nil {
		t.Skip(err)
	}
	st := pres.masterStates["ppt/slideMasters/slideMaster5.xml"]
	if st == nil {
		t.Skip("no master5")
	}
	// re-point the active state at master5's theme the way masterStateFor did
	pres.applyThemeState(st.theme)
	t.Logf("active bg2: %q lt2: %q", pres.themeColors["bg2"], pres.themeColors["lt2"])
	t.Logf("active media: %d entries", len(pres.activeThemeMedia))
	f, ds := resolveThemeBgFillStyle(pres, 3, "bg2", nil)
	if ds != nil {
		t.Logf("texture ds: duotone=%v A=%q B=%q tile=%v data=%d", ds.hasDuotone, ds.duotoneA.ARGB, ds.duotoneB.ARGB, ds.tile, len(ds.data))
	} else if f != nil {
		t.Logf("fill fallback type=%v colour=%q", f.Type, f.Color.ARGB)
	} else {
		t.Logf("both nil")
	}
}
