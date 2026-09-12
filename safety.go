package gopresentation

import (
	"errors"
	"fmt"
	"runtime/debug"
)

// This file implements the library's fail-closed boundary. The reader and the
// renderer walk deeply nested, attacker-controllable XML (OOXML packages from
// arbitrary sources), and a single unguarded nil map, slice index or type
// assertion would otherwise take down the caller's whole process. Every public
// entry point that parses or rasterizes untrusted input therefore recovers
// panics and reports them as an error, so callers do not need to wrap each call
// in their own recover.

// PanicError reports that a panic was recovered inside the library. It is
// returned by the public entry points (Open, Read, GetSlide, SlideToImage,
// SaveSlideAsImage, ...) instead of letting the panic escape.
type PanicError struct {
	// Op names the entry point that recovered, e.g. "Open" or "SlideToImage".
	Op string
	// Value is the recovered panic value.
	Value any
	// Prior is the error already being returned when the panic happened, if
	// any. This happens when a panic occurs during deferred cleanup after the
	// function had decided on its result.
	Prior error
	// Stack is the stack trace captured at recovery time.
	Stack []byte
}

func (e *PanicError) Error() string {
	if e.Prior != nil {
		return fmt.Sprintf("%s: recovered from panic: %v (prior error: %v)", e.Op, e.Value, e.Prior)
	}
	return fmt.Sprintf("%s: recovered from panic: %v", e.Op, e.Value)
}

// Unwrap exposes the recovered value when it is itself an error, and the prior
// error if one was pending, so errors.Is / errors.As can still find them.
func (e *PanicError) Unwrap() []error {
	var errs []error
	if e.Prior != nil {
		errs = append(errs, e.Prior)
	}
	if err, ok := e.Value.(error); ok {
		errs = append(errs, err)
	}
	return errs
}

// StackTrace returns the stack captured when the panic was recovered. The
// trace starts at the entry point; the frames below it were already unwound by
// the time the panic could be recovered.
func (e *PanicError) StackTrace() []byte { return e.Stack }

// recoverToError recovers a panic into *errp. It is meant to be called from a
// deferred statement in a public entry point that has named error results:
//
//	func Open(path string) (pres *Presentation, err error) {
//		defer recoverToError(&err, "Open")
//		...
//	}
//
// When errp or the value it points to is nil the panic is re-raised, because
// silently swallowing it would hide the failure entirely.
func recoverToError(errp *error, op string) {
	r := recover()
	if r == nil {
		return
	}
	if errp == nil {
		panic(r)
	}
	pe := &PanicError{Op: op, Value: r, Prior: *errp, Stack: debug.Stack()}
	*errp = pe
}

// ErrIsPanic reports whether err was produced by recovering a panic inside the
// library, returning the recovered value when it was.
func ErrIsPanic(err error) (any, bool) {
	var pe *PanicError
	if errors.As(err, &pe) {
		return pe.Value, true
	}
	return nil, false
}
