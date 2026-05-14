package forms

import "context"

type FormBackend interface {
	ListForms(ctx context.Context) ([]FormInfo, error)
	InspectForm(ctx context.Context, name string) (*FormSnapshot, error)
}

type DesignerFormBackend interface {
	FormBackend
	SnapshotForm(ctx context.Context, name string) (*FormSpec, error)
}

type RuntimeFormBackend interface {
	FormBackend
	ExportImage(ctx context.Context, name string, outPath string) error
}

type FormInfo struct {
	Name          string `json:"name"`
	ComponentType string `json:"component_type,omitempty"`
	HasFRX        bool   `json:"has_frx,omitempty"`
	SourcePath    string `json:"source_path,omitempty"`
	FRXPath       string `json:"frx_path,omitempty"`
}

type FormSnapshot struct {
	Name             string            `json:"name"`
	Caption          string            `json:"caption,omitempty"`
	Width            float64           `json:"width,omitempty"`
	Height           float64           `json:"height,omitempty"`
	Controls         []ControlSnapshot `json:"controls,omitempty"`
	Basis            string            `json:"basis"`
	CoordinateSystem string            `json:"coordinate_system,omitempty"`
	Warnings         []FormWarning     `json:"warnings,omitempty"`
}

type ControlSnapshot struct {
	Name          string            `json:"name"`
	Type          string            `json:"type"`
	ProgID        string            `json:"prog_id,omitempty"`
	Caption       *string           `json:"caption,omitempty"`
	Text          *string           `json:"text,omitempty"`
	Value         any               `json:"value,omitempty"`
	Left          float64           `json:"left,omitempty"`
	Top           float64           `json:"top,omitempty"`
	Width         float64           `json:"width,omitempty"`
	Height        float64           `json:"height,omitempty"`
	TabIndex      *int              `json:"tab_index,omitempty"`
	SelectedIndex *int              `json:"selected_index,omitempty"`
	Enabled       *bool             `json:"enabled,omitempty"`
	Visible       *bool             `json:"visible,omitempty"`
	List          []string          `json:"list,omitempty"`
	Properties    map[string]any    `json:"properties,omitempty"`
	Unsupported   []string          `json:"unsupported,omitempty"`
	Controls      []ControlSnapshot `json:"controls,omitempty"`
}

type FormWarning struct {
	Code    string `json:"code"`
	Message string `json:"message"`
	Control string `json:"control,omitempty"`
}
