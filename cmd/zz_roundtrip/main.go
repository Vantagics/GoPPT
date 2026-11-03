// Command zz_roundtrip reads a deck with GoPPT and writes it back out, so a
// generation-side fidelity round can put PowerPoint's export of the rewritten
// file next to its export of the original: whatever the writer loses or
// distorts shows up as a pixel difference between those two renders.
//
//	zz_roundtrip <in.pptx> <out.pptx>
package main

import (
	"fmt"
	"os"

	gopresentation "github.com/Vantagics/GoPPT"
)

func main() {
	if len(os.Args) != 3 {
		fmt.Fprintln(os.Stderr, "usage: zz_roundtrip <in.pptx> <out.pptx>")
		os.Exit(2)
	}
	in, out := os.Args[1], os.Args[2]

	pres, err := gopresentation.Open(in)
	if err != nil {
		fmt.Fprintf(os.Stderr, "read %s: %v\n", in, err)
		os.Exit(1)
	}
	if err := pres.Save(out); err != nil {
		fmt.Fprintf(os.Stderr, "write %s: %v\n", out, err)
		os.Exit(1)
	}
}
