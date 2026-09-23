package main

import (
	"bytes"
	"fmt"
	"os"

	goppt "github.com/Vantagics/GoPPT"
)

func main() {
	p := goppt.New()
	var buf bytes.Buffer
	if err := p.WriteTo(&buf); err != nil {
		panic(err)
	}
	_ = os.WriteFile(os.Args[1], buf.Bytes(), 0644)
	fmt.Println("written", buf.Len())
}
