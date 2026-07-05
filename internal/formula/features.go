package formula

const (
	FeatureExternalReference    = "external_reference"
	Feature3DReference          = "3d_reference"
	FeatureStructuredReference  = "structured_reference"
	FeatureSpillReference       = "spill_reference"
	FeatureImplicitIntersection = "implicit_intersection"
)

type Feature struct {
	Code string `json:"code"`
	Raw  string `json:"raw"`
	Span Span   `json:"span"`
}
