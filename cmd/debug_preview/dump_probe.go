package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// dumpPart prints one part of a .pptx (or searches it) so the XML that the
// reader consumes can be inspected without a working unzip tool.
//
// Output is ASCII-only (strconv.Quote), so the result is unaffected by whatever
// codepage the log pipeline assumes. A CJK string shows up as \u7ec4\u7ec7...
// which is the only unambiguous way to tell "the file is UTF-8" from "the file
// was mis-decoded before it got here".
//
//	go run ./cmd/debug_preview --part <file.pptx> <part-name> [substring]
func dumpPart(path, part, only string) {
	zr, err := zip.OpenReader(path)
	if err != nil {
		fmt.Printf("open: %v\n", err)
		return
	}
	defer zr.Close()
	for _, f := range zr.File {
		if f.Name != part {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			fmt.Printf("open part: %v\n", err)
			return
		}
		data, _ := io.ReadAll(rc)
		rc.Close()
		s := string(data)
		fmt.Printf("part %s: %d bytes\n", part, len(data))
		if only == "" {
			fmt.Println(strconv.Quote(s))
			return
		}
		n := 0
		for i := 0; ; {
			j := strings.Index(s[i:], only)
			if j < 0 {
				break
			}
			abs := i + j
			lo := abs - 120
			if lo < 0 {
				lo = 0
			}
			hi := abs + len(only) + 120
			if hi > len(s) {
				hi = len(s)
			}
			n++
			fmt.Printf("--- hit %d ---\n%s\n\n", n, strconv.Quote(s[lo:hi]))
			i = abs + len(only)
		}
		fmt.Printf("hits: %d\n", n)
		return
	}
	fmt.Printf("part %q not found\n", part)
}

// dumpPartPretty prints one part with a line break before every '<' so the XML
// can be read top to bottom instead of as one enormous line.
//
//	strconv.Quote is applied per line, so the output stays ASCII-only for the
//	same reason dumpPart does it: a CJK run then shows as \u7ec4\u7ec7..., which
//	distinguishes "the part is UTF-8" from "the part was mis-decoded upstream".
//
//	go run ./cmd/debug_preview --pretty <file.pptx> <part-name>
func dumpPartPretty(path, part string) {
	zr, err := zip.OpenReader(path)
	if err != nil {
		fmt.Printf("open: %v\n", err)
		return
	}
	defer zr.Close()
	for _, f := range zr.File {
		if f.Name != part {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			fmt.Printf("open part: %v\n", err)
			return
		}
		data, _ := io.ReadAll(rc)
		rc.Close()
		fmt.Printf("part %s: %d bytes\n", part, len(data))
		s := strings.ReplaceAll(string(data), "><", ">\n<")
		for i, line := range strings.Split(s, "\n") {
			fmt.Printf("%4d %s\n", i+1, strconv.Quote(line))
		}
		return
	}
	fmt.Printf("part %q not found\n", part)
}
