// Command zz_probe48 builds probe decks for the slide47 title defect (whole
// 2-line autofit block renders +18px lower than COM while line pitch and glyph
// size match exactly). Two decks:
//
//	out_deck/r48bisect.pptx — 4 slides bisecting WHICH property makes our
//	reader drop the spliced shape: s1 control (solidFill rect, empty body,
//	r68b pattern) / s2 +txBox="1" / s3 noFill / s4 full text body.
//
//	out_deck/r69probe.pptx — 6-slide autofit ladder (same geometry as the
//	slide47 title box, anchor=ctr): A fontScale+lnSpcReduction 2 paras /
//	B fontScale only / C lnSpcReduction only / D both levers 1 para /
//	E no autofit 2 paras / F no autofit 1 para.
package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"strings"

	pres "github.com/Vantagics/GoPPT"
)

const body = `<a:pPr algn="ctr"><a:spcAft><a:spcPts val="0"/></a:spcAft></a:pPr><a:r><a:rPr lang="en-US" sz="4400" dirty="0"/><a:t>%s</a:t></a:r>`

const slideTmpl = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="T"/>%s<p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="685800" y="425450"/><a:ext cx="7772400" cy="987425"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        %s
      </p:spPr>
      <p:txBody><a:bodyPr %s>%s</a:bodyPr><a:lstStyle/>%s</p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`

var (
	redFill = `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`
	noFill  = `<a:noFill/>`
	txBox   = `<p:cNvSpPr txBox="1"/>`
	plainNV = `<p:cNvSpPr/>`

	fitBoth  = `<a:normAutofit fontScale="77500" lnSpcReduction="20000"/>`
	fitScale = `<a:normAutofit fontScale="77500"/>`
	fitLn    = `<a:normAutofit lnSpcReduction="20000"/>`
)

func slide(nvSpPr, fill, bodyPrAttrs, autofit, txBody string) string {
	return fmt.Sprintf(slideTmpl, nvSpPr, fill, bodyPrAttrs, autofit, txBody)
}

func paras(texts ...string) string {
	var sb strings.Builder
	for _, t := range texts {
		sb.WriteString("<a:p>")
		fmt.Fprintf(&sb, body, t)
		sb.WriteString("</a:p>")
	}
	return sb.String()
}

func build(out string, slides map[string]string, n int) {
	p := pres.New()
	for i := 1; i < n; i++ {
		p.CreateSlide()
	}
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
	spliced := 0
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}
		data, _ := io.ReadAll(rc)
		rc.Close()
		if repl, ok := slides[f.Name]; ok {
			data = []byte(repl)
			spliced++
		}
		w, err := zw.Create(f.Name)
		if err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}
		w.Write(data)
	}
	zw.Close()
	if spliced != len(slides) {
		fmt.Fprintf(os.Stderr, "FATAL: spliced %d of %d slides — map keys must be full zip paths\n", spliced, len(slides))
		os.Exit(1)
	}
	if err := os.WriteFile(out, outBuf.Bytes(), 0o644); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	fmt.Println("wrote", out, "spliced", spliced)
}

func main() {
	bisectOut := "out_deck/r48bisect.pptx"
	probeOut := "out_deck/r69probe.pptx"
	if len(os.Args) > 1 {
		bisectOut = os.Args[1]
	}
	if len(os.Args) > 2 {
		probeOut = os.Args[2]
	}
	empty := `<a:p/>`
	bisect := map[string]string{
		"ppt/slides/slide1.xml": slide(plainNV, redFill, ``, ``, empty),                                 // control (r68b pattern)
		"ppt/slides/slide2.xml": slide(txBox, redFill, ``, ``, empty),                                   // +txBox
		"ppt/slides/slide3.xml": slide(plainNV, noFill, ``, ``, empty),                                  // +noFill
		"ppt/slides/slide4.xml": slide(txBox, noFill, `anchor="ctr"`, ``, paras("Tasks for Districts")), // full
	}
	build(bisectOut, bisect, 4)

	ladder := map[string]string{
		"ppt/slides/slide1.xml": slide(txBox, noFill, `anchor="ctr"`, fitBoth, paras("Preparation for New Standards", "Tasks for Districts")),
		"ppt/slides/slide2.xml": slide(txBox, noFill, `anchor="ctr"`, fitScale, paras("Preparation for New Standards", "Tasks for Districts")),
		"ppt/slides/slide3.xml": slide(txBox, noFill, `anchor="ctr"`, fitLn, paras("Preparation for New Standards", "Tasks for Districts")),
		"ppt/slides/slide4.xml": slide(txBox, noFill, `anchor="ctr"`, fitBoth, paras("Tasks for Districts")),
		"ppt/slides/slide5.xml": slide(txBox, noFill, `anchor="ctr"`, ``, paras("Preparation for New Standards", "Tasks for Districts")),
		"ppt/slides/slide6.xml": slide(txBox, noFill, `anchor="ctr"`, ``, paras("Tasks for Districts")),
		// G/H: single-line lever isolation (44pt "Tasks for Districts" never wraps)
		"ppt/slides/slide7.xml": slide(txBox, noFill, `anchor="ctr"`, fitLn, paras("Tasks for Districts")),
		"ppt/slides/slide8.xml": slide(txBox, noFill, `anchor="ctr"`, fitScale, paras("Tasks for Districts")),
	}
	build(probeOut, ladder, 8)
}
