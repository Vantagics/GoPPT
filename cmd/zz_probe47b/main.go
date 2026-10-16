// Command zz_probe47b builds a single-rect deck replicating drawing10's last
// SmartArt column (the peach one): schemeClr accent5 + tint 40% + alpha 90% +
// hueOff -10740482 + satOff 48253 + lumOff 3317. Rendering it ours and COM and
// sampling the centre pins whether the r68 hueOff/satOff/lumOff model matches
// PowerPoint's own reading of the exact transform combo slide47 uses.
package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"

	pres "github.com/Vantagics/GoPPT"
)

const colLast = `<a:solidFill><a:schemeClr val="accent5"><a:tint val="40000"/><a:alpha val="90000"/><a:hueOff val="-10740482"/><a:satOff val="48253"/><a:lumOff val="3317"/><a:alphaOff val="0"/></a:schemeClr></a:solidFill>`

const slide = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="C"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="200000" y="200000"/><a:ext cx="4000000" cy="4000000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
` + colLast + `
      </p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`

func main() {
	out := "out_deck/r68b.pptx"
	if len(os.Args) > 1 {
		out = os.Args[1]
	}
	p := pres.New()
	var buf bytes.Buffer
	if err := p.WriteTo(&buf); err != nil {
		fmt.Fprintln(os.Stderr, "write:", err)
		os.Exit(1)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		fmt.Fprintln(os.Stderr, "re-open:", err)
		os.Exit(1)
	}
	var outBuf bytes.Buffer
	zw := zip.NewWriter(&outBuf)
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}
		data, _ := io.ReadAll(rc)
		rc.Close()
		if f.Name == "ppt/slides/slide1.xml" {
			data = []byte(slide)
		}
		w, err := zw.Create(f.Name)
		if err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}
		w.Write(data)
	}
	zw.Close()
	if err := os.WriteFile(out, outBuf.Bytes(), 0o644); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	fmt.Println("wrote", out)
}
