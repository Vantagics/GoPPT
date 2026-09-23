package main

// r42 probe: parse 00022823 chart1.xml through the reader, print series marker/outline state.
import (
	"fmt"
	"os"

	goppt "github.com/Vantagics/GoPPT"
)

func main() {
	deck, err := goppt.Open(os.Args[1])
	if err != nil {
		fmt.Println("open:", err)
		return
	}
	sl, err := deck.GetSlide(4)
	if err != nil {
		fmt.Println("slide:", err)
		return
	}
	for _, sh := range sl.GetShapes() {
		cs, ok := sh.(*goppt.ChartShape)
		if !ok {
			continue
		}
		lc, ok := cs.GetPlotArea().GetType().(*goppt.LineChart)
		if !ok {
			continue
		}
		for si, ser := range lc.Series {
			mstr := "<nil>"
			if m := ser.Marker; m != nil {
				mstr = fmt.Sprintf("{sym=%q size=%d}", m.Symbol, m.Size)
			}
			ostr := "<nil>"
			if o := ser.Outline; o != nil {
				ostr = fmt.Sprintf("{w=%d col=%v}", o.Width, o.Color)
			}
			fmt.Printf("ser idx=%d marker=%s outline=%s dash=%q\n", si, mstr, ostr, ser.LineDash)
		}
	}
}
