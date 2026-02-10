package gopresentation

import "fmt"

// Version information for the GoPresentation library.
const (
	VersionMajor = 1
	VersionMinor = 0
	VersionPatch = 0
)

// Version is the full version string of the GoPresentation library.
var Version = fmt.Sprintf("%d.%d.%d", VersionMajor, VersionMinor, VersionPatch)
