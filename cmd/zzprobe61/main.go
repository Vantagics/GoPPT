package main

import (
	goppt "github.com/Vantagics/GoPPT"
)

func main() {
	for _, d := range []struct{ deck, out string }{
		{`D:/workprj/officeread/testdata/web-samples/samples/pptx/00022693.pptx`, `out_deck/r61/go22693/slide%02d.png`},
		{`D:/workprj/officeread/testdata/web-samples/samples/pptx/00022823.pptx`, `out_deck/r61/go22823/slide%02d.png`},
	} {
		pres, err := goppt.Open(d.deck)
		if err != nil {
			panic(err)
		}
		opts := goppt.DefaultRenderOptions()
		opts.Width = 1600
		opts.FontCache = goppt.NewFontCache()
		if err := pres.SaveSlidesAsImages(d.out, opts); err != nil {
			panic(err)
		}
	}
}
