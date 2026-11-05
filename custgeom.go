package gopresentation

import (
	"math"
	"strconv"
	"strings"
)

// custGuide is one <a:gd> entry of a custGeom's <a:gdLst>: a named formula the
// path's point attributes reference. PowerPoint converts freeforms into this
// form routinely (every "Freeform N" of a pasted SmartArt or a chart pasted as
// picture carries one), so a path whose <a:pt x="f5"/> never resolves renders
// as a degenerate empty polygon.
type custGuide struct {
	name string
	fmla string
}

// evalCustGeomGuides evaluates a gdLst into a name → value map. w/h are the
// path's coordinate space; "ss" is the shortest of the two, the third built-in
// the spec defines alongside them. Guides are resolved lazily with
// memoisation so declaration order never matters; a reference cycle resolves
// to 0 (the spec forbids cycles, so this only guards against them).
func evalCustGeomGuides(guides []custGuide, w, h int64) map[string]float64 {
	vals := map[string]float64{"w": float64(w), "h": float64(h)}
	if w < h {
		vals["ss"] = float64(w)
	} else {
		vals["ss"] = float64(h)
	}
	if len(guides) == 0 {
		return vals
	}
	byName := make(map[string]string, len(guides))
	for _, g := range guides {
		if _, dup := byName[g.name]; !dup {
			byName[g.name] = g.fmla
		}
	}
	var eval func(name string) float64
	stack := map[string]bool{}
	eval = func(name string) float64 {
		if v, ok := vals[name]; ok {
			return v
		}
		fmla, ok := byName[name]
		if !ok || stack[name] {
			return 0
		}
		stack[name] = true
		v := evalCustGeomFmla(fmla, eval)
		delete(stack, name)
		vals[name] = v
		return v
	}
	for name := range byName {
		eval(name)
	}
	return vals
}

// evalCustGeomFmla evaluates one guide formula: the first token is the
// operator, the rest its arguments. An argument is either a number literal or
// the name of another guide (or w/h/ss). Angles are in 60000ths of a degree,
// matching how the path's arcTo stores them.
func evalCustGeomFmla(fmla string, eval func(string) float64) float64 {
	toks := strings.Fields(fmla)
	if len(toks) == 0 {
		return 0
	}
	arg := func(i int) float64 {
		if i >= len(toks) {
			return 0
		}
		if v, err := strconv.ParseFloat(toks[i], 64); err == nil {
			return v
		}
		return eval(toks[i])
	}
	switch toks[0] {
	case "val":
		return arg(1)
	case "*/":
		// a*b/c; a zero divisor yields 0 rather than Inf (a mid-round guide
		// can legally divide by a guide that evaluated to 0).
		if len(toks) < 4 {
			return 0
		}
		d := arg(3)
		if d == 0 {
			return 0
		}
		return arg(1) * arg(2) / d
	case "+-":
		if len(toks) < 4 {
			return 0
		}
		return arg(1) + arg(2) - arg(3)
	case "?:":
		if len(toks) < 4 {
			return 0
		}
		if arg(1) != 0 {
			return arg(2)
		}
		return arg(3)
	case "pin":
		// clamp(x, lo, hi) — the spec's three-argument bound.
		if len(toks) < 4 {
			return 0
		}
		x, lo, hi := arg(1), arg(2), arg(3)
		if x < lo {
			return lo
		}
		if x > hi {
			return hi
		}
		return x
	case "min":
		v := arg(1)
		for i := 2; i < len(toks); i++ {
			if a := arg(i); a < v {
				v = a
			}
		}
		return v
	case "max":
		v := arg(1)
		for i := 2; i < len(toks); i++ {
			if a := arg(i); a > v {
				v = a
			}
		}
		return v
	case "abs":
		return math.Abs(arg(1))
	case "sqrt":
		return math.Sqrt(math.Abs(arg(1)))
	case "at2":
		// atan2(a, b) expressed back in 60000ths of a degree.
		return math.Atan2(arg(1), arg(2)) * 180 / math.Pi * 60000
	case "sin":
		return math.Sin(arg(1)/60000*math.Pi/180) * arg(2)
	case "cos":
		return math.Cos(arg(1)/60000*math.Pi/180) * arg(2)
	case "tan":
		return math.Tan(arg(1)/60000*math.Pi/180) * arg(2)
	}
	// A single-token formula ("ss", or a bare number) is not an operator at
	// all: value = that expression.
	if v, err := strconv.ParseFloat(toks[0], 64); err == nil {
		return v
	}
	return eval(toks[0])
}
