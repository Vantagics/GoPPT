package gopresentation

import (
	"testing"
)

func TestProbeTexStyles(t *testing.T) {
	pres, err := Open("F:/pptx_corpus/web/themes.pptx")
	if err != nil {
		t.Skip(err)
	}
	st := pres.masterStates["ppt/slideMasters/slideMaster5.xml"]
	if st == nil || st.theme == nil {
		t.Skip("master5 theme missing")
	}
	for i, s := range st.theme.bgFillStyles {
		t.Logf("bg[%d]: blip=%v rel=%q tile=%v duo=%d", i, s.blip, s.blipRelID, s.tile, len(s.blipDuo))
		for j, d := range s.blipDuo {
			t.Logf("   duo[%d] ops=%v", j, d)
		}
	}
	if st.theme.media != nil {
		for k, m := range st.theme.media {
			t.Logf("media[%s] = %d bytes ext=%s", k, len(m.data), m.ext)
		}
	} else {
		t.Logf("media map NIL")
	}
	t.Logf("bg2 colour: %q", st.theme.colors["bg2"])
	t.Logf("lt2 colour: %q", st.theme.colors["lt2"])
}
