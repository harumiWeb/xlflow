package oforms

type fieldKind uint8

const (
	fieldUnsigned fieldKind = iota
	fieldSigned
	fieldStringLength
	fieldMarker
)

type extraKind uint8

const (
	extraSize extraKind = iota
	extraPosition
	extraString
	extraArray
)

type dataField struct {
	bit  uint8
	name string
	size int
	kind fieldKind
}

type extraField struct {
	bit      uint8
	name     string
	kind     extraKind
	sizeFrom string
}

type recordSpec struct {
	typeName       string
	major          uint8
	mask64         bool
	stopAfterExtra bool
	textProps      bool
	rawTail        bool
	data           []dataField
	extra          []extraField
	stream         []streamField
	flags          map[uint8]flagField
}

type streamField struct {
	bit  uint8
	name string
}

type flagField struct {
	name  string
	value int64
}

var textPropsSpec = recordSpec{
	typeName: "TextProps", major: 2, stopAfterExtra: true,
	data: []dataField{
		{0, "FontName", 4, fieldStringLength}, {1, "FontEffects", 4, fieldUnsigned},
		{2, "FontHeight", 4, fieldUnsigned}, {4, "FontCharSet", 1, fieldUnsigned},
		{5, "FontPitchAndFamily", 1, fieldUnsigned}, {6, "ParagraphAlign", 1, fieldUnsigned},
		{7, "FontWeight", 2, fieldUnsigned},
	},
	extra: []extraField{{0, "FontName", extraString, ""}},
}

var morphDataSpec = recordSpec{
	typeName: "MorphData", major: 2, mask64: true, textProps: true, rawTail: true,
	data: []dataField{
		{0, "VariousPropertyBits", 4, fieldUnsigned}, {1, "BackColor", 4, fieldUnsigned},
		{2, "ForeColor", 4, fieldUnsigned}, {3, "MaxLength", 4, fieldUnsigned},
		{4, "BorderStyle", 1, fieldUnsigned}, {5, "ScrollBars", 1, fieldUnsigned},
		{6, "DisplayStyle", 1, fieldUnsigned}, {7, "MousePointer", 1, fieldUnsigned},
		{9, "PasswordChar", 2, fieldUnsigned}, {10, "ListWidth", 4, fieldUnsigned},
		{11, "BoundColumn", 2, fieldUnsigned}, {12, "TextColumn", 2, fieldSigned},
		{13, "ColumnCount", 2, fieldSigned}, {14, "ListRows", 2, fieldUnsigned},
		{15, "cColumnInfo", 2, fieldUnsigned}, {16, "MatchEntry", 1, fieldUnsigned},
		{17, "ListStyle", 1, fieldUnsigned}, {18, "ShowDropButtonWhen", 1, fieldUnsigned},
		{20, "DropButtonStyle", 1, fieldUnsigned}, {21, "MultiSelect", 1, fieldUnsigned},
		{22, "Value", 4, fieldStringLength}, {23, "Caption", 4, fieldStringLength},
		{24, "PicturePosition", 4, fieldUnsigned}, {25, "BorderColor", 4, fieldUnsigned},
		{26, "SpecialEffect", 4, fieldUnsigned}, {27, "MouseIcon", 2, fieldMarker},
		{28, "Picture", 2, fieldMarker}, {29, "Accelerator", 2, fieldUnsigned},
		{32, "GroupName", 4, fieldStringLength},
	},
	extra: []extraField{
		{8, "Size", extraSize, ""}, {22, "Value", extraString, ""},
		{23, "Caption", extraString, ""}, {32, "GroupName", extraString, ""},
	},
	stream: []streamField{{27, "MouseIcon"}, {28, "Picture"}},
}

var commandButtonSpec = recordSpec{
	typeName: "CommandButton", major: 2, textProps: true,
	data: []dataField{
		{0, "ForeColor", 4, fieldUnsigned}, {1, "BackColor", 4, fieldUnsigned},
		{2, "VariousPropertyBits", 4, fieldUnsigned}, {3, "Caption", 4, fieldStringLength},
		{4, "PicturePosition", 4, fieldUnsigned}, {6, "MousePointer", 1, fieldUnsigned},
		{7, "Picture", 2, fieldMarker}, {8, "Accelerator", 2, fieldUnsigned},
		{10, "MouseIcon", 2, fieldMarker},
	},
	extra:  []extraField{{3, "Caption", extraString, ""}, {5, "Size", extraSize, ""}},
	stream: []streamField{{7, "Picture"}, {10, "MouseIcon"}},
	flags:  map[uint8]flagField{9: {name: "TakeFocusOnClick", value: 0}},
}

var labelSpec = recordSpec{
	typeName: "Label", major: 2, textProps: true,
	data: []dataField{
		{0, "ForeColor", 4, fieldUnsigned}, {1, "BackColor", 4, fieldUnsigned},
		{2, "VariousPropertyBits", 4, fieldUnsigned}, {3, "Caption", 4, fieldStringLength},
		{4, "PicturePosition", 4, fieldUnsigned}, {6, "MousePointer", 1, fieldUnsigned},
		{7, "BorderColor", 4, fieldUnsigned}, {8, "BorderStyle", 2, fieldUnsigned},
		{9, "SpecialEffect", 2, fieldUnsigned}, {10, "Picture", 2, fieldMarker},
		{11, "Accelerator", 2, fieldUnsigned}, {12, "MouseIcon", 2, fieldMarker},
	},
	extra:  []extraField{{3, "Caption", extraString, ""}, {5, "Size", extraSize, ""}},
	stream: []streamField{{10, "Picture"}, {12, "MouseIcon"}},
}

var imageSpec = recordSpec{
	typeName: "Image", major: 2,
	data: []dataField{
		{3, "BorderColor", 4, fieldUnsigned}, {4, "BackColor", 4, fieldUnsigned},
		{5, "BorderStyle", 1, fieldUnsigned}, {6, "MousePointer", 1, fieldUnsigned},
		{7, "PictureSizeMode", 1, fieldUnsigned}, {8, "SpecialEffect", 1, fieldUnsigned},
		{10, "Picture", 2, fieldMarker}, {11, "PictureAlignment", 1, fieldUnsigned},
		{13, "VariousPropertyBits", 4, fieldUnsigned}, {14, "MouseIcon", 2, fieldMarker},
	},
	extra:  []extraField{{9, "Size", extraSize, ""}},
	stream: []streamField{{10, "Picture"}, {14, "MouseIcon"}},
}

var spinButtonSpec = recordSpec{
	typeName: "SpinButton", major: 2,
	data: []dataField{
		{0, "ForeColor", 4, fieldUnsigned}, {1, "BackColor", 4, fieldUnsigned},
		{2, "VariousPropertyBits", 4, fieldUnsigned}, {5, "Min", 4, fieldSigned},
		{6, "Max", 4, fieldSigned}, {7, "Position", 4, fieldSigned},
		{8, "PrevEnabled", 4, fieldSigned}, {9, "NextEnabled", 4, fieldSigned},
		{10, "SmallChange", 4, fieldSigned}, {11, "Orientation", 4, fieldSigned},
		{12, "Delay", 4, fieldUnsigned}, {13, "MouseIcon", 2, fieldMarker},
		{14, "MousePointer", 1, fieldUnsigned},
	},
	extra:  []extraField{{3, "Size", extraSize, ""}},
	stream: []streamField{{13, "MouseIcon"}},
}

var scrollBarSpec = recordSpec{
	typeName: "ScrollBar", major: 2,
	data: []dataField{
		{0, "ForeColor", 4, fieldUnsigned}, {1, "BackColor", 4, fieldUnsigned},
		{2, "VariousPropertyBits", 4, fieldUnsigned}, {4, "MousePointer", 1, fieldUnsigned},
		{5, "Min", 4, fieldSigned}, {6, "Max", 4, fieldSigned},
		{7, "Position", 4, fieldSigned}, {9, "PrevEnabled", 4, fieldSigned},
		{10, "NextEnabled", 4, fieldSigned}, {11, "SmallChange", 4, fieldSigned},
		{12, "LargeChange", 4, fieldSigned}, {13, "Orientation", 4, fieldSigned},
		{14, "ProportionalThumb", 2, fieldSigned}, {15, "Delay", 4, fieldUnsigned},
		{16, "MouseIcon", 2, fieldMarker},
	},
	extra:  []extraField{{3, "Size", extraSize, ""}},
	stream: []streamField{{16, "MouseIcon"}},
}

var tabStripSpec = recordSpec{
	typeName: "TabStrip", major: 2, textProps: true, rawTail: true,
	flags: map[uint8]flagField{10: {name: "MultiRow", value: 1}, 13: {name: "Tooltips", value: 0}, 19: {name: "NewVersion", value: 1}},
	data: []dataField{
		{0, "ListIndex", 4, fieldSigned}, {1, "BackColor", 4, fieldUnsigned},
		{2, "ForeColor", 4, fieldUnsigned}, {5, "ItemsSize", 4, fieldUnsigned},
		{6, "MousePointer", 1, fieldUnsigned}, {8, "TabOrientation", 4, fieldUnsigned},
		{9, "TabStyle", 4, fieldUnsigned}, {11, "TabFixedWidth", 4, fieldUnsigned},
		{12, "TabFixedHeight", 4, fieldUnsigned}, {15, "TipStringsSize", 4, fieldUnsigned},
		{17, "NamesSize", 4, fieldUnsigned}, {18, "VariousPropertyBits", 4, fieldUnsigned},
		{20, "TabsAllocated", 4, fieldUnsigned}, {21, "TagsSize", 4, fieldUnsigned},
		{22, "TabData", 4, fieldUnsigned}, {23, "AcceleratorsSize", 4, fieldUnsigned},
		{24, "MouseIcon", 2, fieldMarker},
	},
	extra: []extraField{
		{4, "Size", extraSize, ""}, {5, "Items", extraArray, "ItemsSize"},
		{15, "TipStrings", extraArray, "TipStringsSize"}, {17, "TabNames", extraArray, "NamesSize"},
		{21, "Tags", extraArray, "TagsSize"}, {23, "Accelerators", extraArray, "AcceleratorsSize"},
	},
	stream: []streamField{{24, "MouseIcon"}},
}

var formSpec = recordSpec{
	typeName: "Form", major: 4, stopAfterExtra: true,
	data: []dataField{
		{1, "BackColor", 4, fieldUnsigned}, {2, "ForeColor", 4, fieldUnsigned},
		{3, "NextAvailableID", 4, fieldUnsigned}, {6, "BooleanProperties", 4, fieldUnsigned},
		{7, "BorderStyle", 1, fieldUnsigned}, {8, "MousePointer", 1, fieldUnsigned},
		{9, "ScrollBars", 1, fieldUnsigned}, {13, "GroupCnt", 4, fieldSigned},
		{15, "MouseIcon", 2, fieldMarker}, {16, "Cycle", 1, fieldUnsigned},
		{17, "SpecialEffect", 1, fieldUnsigned}, {18, "BorderColor", 4, fieldUnsigned},
		{19, "Caption", 4, fieldStringLength}, {20, "Font", 2, fieldMarker},
		{21, "Picture", 2, fieldMarker}, {22, "Zoom", 4, fieldUnsigned},
		{23, "PictureAlignment", 1, fieldUnsigned}, {25, "PictureSizeMode", 1, fieldUnsigned},
		{26, "ShapeCookie", 4, fieldUnsigned}, {27, "DrawBuffer", 4, fieldUnsigned},
	},
	extra: []extraField{
		{10, "DisplayedSize", extraSize, ""}, {11, "LogicalSize", extraSize, ""},
		{12, "ScrollPosition", extraPosition, ""}, {19, "Caption", extraString, ""},
	},
}

var specsByCacheIndex = map[uint16]*recordSpec{
	12: &imageSpec, 15: &morphDataSpec, 16: &spinButtonSpec,
	17: &commandButtonSpec, 18: &tabStripSpec, 21: &labelSpec,
	23: &morphDataSpec, 24: &morphDataSpec, 25: &morphDataSpec,
	26: &morphDataSpec, 27: &morphDataSpec, 28: &morphDataSpec,
	47: &scrollBarSpec,
}

var controlKinds = map[uint16]string{
	7: "MSForms.Form", 12: "MSForms.Image", 14: "MSForms.Frame",
	15: "MSForms.MorphData", 16: "MSForms.SpinButton", 17: "MSForms.CommandButton",
	18: "MSForms.TabStrip", 21: "MSForms.Label", 23: "MSForms.TextBox",
	24: "MSForms.ListBox", 25: "MSForms.ComboBox", 26: "MSForms.CheckBox",
	27: "MSForms.OptionButton", 28: "MSForms.ToggleButton", 47: "MSForms.ScrollBar",
	57: "MSForms.MultiPage",
}

var containerCacheIndices = map[uint16]struct{}{7: {}, 14: {}, 57: {}}
