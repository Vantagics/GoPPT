package gopresentation

import (
	"testing"
)

func TestProbeKey02Sizes(t *testing.T) {
	pres, err := Open("F:/pptx_corpus/web/KEY02.pptx")
	if err != nil {
		t.Skip(err)
	}
	for _, sh := range pres.slides[0].shapes {
		if ph, ok := sh.(*PlaceholderShape); ok {
			for pi, para := range ph.paragraphs {
				for _, elem := range para.elements {
					if tr, ok := elem.(*TextRun); ok {
						fs := "nil"
						sz := 0
						szSet := false
						if tr.font != nil {
							fs = tr.font.Name
							sz = tr.font.Size
							szSet = tr.font.runSizeSet
						}
						t.Logf("ph=%s para%d run sz=%d set=%v font=%q text=%.30q", ph.phType, pi, sz, szSet, fs, tr.text)
					}
				}
			}
		}
	}
}
