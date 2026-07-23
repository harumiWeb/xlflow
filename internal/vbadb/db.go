package vbadb

import (
	"embed"
	"encoding/json"
	"fmt"
	"io/fs"
	"os"
	"path/filepath"
	"sort"
	"strings"
)

const SchemaVersion = 1

//go:embed builtin/*.json
var builtinFS embed.FS

type DB struct {
	Types        map[string]TypeInfo
	Aliases      map[string]string
	Constants    map[string]ConstantInfo
	ProgIDs      map[string]string
	ProgIDNames  map[string]string
	GlobalValues map[string]string
	GlobalNames  map[string]string
}

type TypeInfo struct {
	Name              string       `json:"name"`
	Library           string       `json:"library,omitempty"`
	Kind              string       `json:"kind,omitempty"`
	Summary           string       `json:"summary,omitempty"`
	ElementType       string       `json:"element_type,omitempty"`
	DefaultMember     string       `json:"default_member,omitempty"`
	DefaultMemberType string       `json:"default_member_type,omitempty"`
	Collection        bool         `json:"collection,omitempty"`
	Confidence        string       `json:"confidence,omitempty"`
	Source            string       `json:"source,omitempty"`
	Aliases           []string     `json:"aliases,omitempty"`
	AssignableTo      []string     `json:"assignable_to,omitempty"`
	Properties        []MemberInfo `json:"properties,omitempty"`
	Methods           []MemberInfo `json:"methods,omitempty"`
	Events            []MemberInfo `json:"events,omitempty"`
}

type MemberInfo struct {
	Name       string      `json:"name"`
	Summary    string      `json:"summary,omitempty"`
	ReturnType string      `json:"return_type,omitempty"`
	Parameters []ParamInfo `json:"parameters,omitempty"`
	ReadOnly   bool        `json:"read_only,omitempty"`
	WriteOnly  bool        `json:"write_only,omitempty"`
	Default    bool        `json:"default,omitempty"`
}

type ParamInfo struct {
	Name       string `json:"name"`
	Type       string `json:"type,omitempty"`
	Optional   bool   `json:"optional,omitempty"`
	ParamArray bool   `json:"param_array,omitempty"`
}

type ConstantInfo struct {
	Name      string `json:"name"`
	Library   string `json:"library,omitempty"`
	Kind      string `json:"kind,omitempty"`
	Type      string `json:"type,omitempty"`
	Value     string `json:"value,omitempty"`
	EnumGroup string `json:"enum_group,omitempty"`
	Summary   string `json:"summary,omitempty"`
}

type fileData struct {
	Types        []TypeInfo        `json:"types,omitempty"`
	Constants    []ConstantInfo    `json:"constants,omitempty"`
	ProgIDs      map[string]string `json:"progids,omitempty"`
	GlobalValues map[string]string `json:"global_values,omitempty"`
}

func LoadBuiltin() (*DB, error) {
	db := New()
	if err := LoadBuiltinInto(db); err != nil {
		return nil, err
	}
	return db, nil
}

func LoadBuiltinInto(db *DB) error {
	if db == nil {
		return fmt.Errorf("vbadb: nil DB")
	}
	err := fs.WalkDir(builtinFS, "builtin", func(path string, d fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if d.IsDir() || strings.ToLower(filepath.Ext(path)) != ".json" {
			return nil
		}
		body, err := builtinFS.ReadFile(path)
		if err != nil {
			return err
		}
		if err := db.MergeJSON(body); err != nil {
			return fmt.Errorf("%s: %w", path, err)
		}
		return nil
	})
	return err
}

func New() *DB {
	return &DB{
		Types:        map[string]TypeInfo{},
		Aliases:      map[string]string{},
		Constants:    map[string]ConstantInfo{},
		ProgIDs:      map[string]string{},
		ProgIDNames:  map[string]string{},
		GlobalValues: map[string]string{},
		GlobalNames:  map[string]string{},
	}
}

func LoadFiles(paths ...string) (*DB, error) {
	db := New()
	for _, path := range paths {
		body, err := os.ReadFile(path)
		if err != nil {
			return nil, err
		}
		if err := db.MergeJSON(body); err != nil {
			return nil, fmt.Errorf("%s: %w", path, err)
		}
	}
	return db, nil
}

func LoadDir(dir string) (*DB, error) {
	db := New()
	err := filepath.WalkDir(dir, func(path string, d fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if d.IsDir() || strings.ToLower(filepath.Ext(path)) != ".json" {
			return nil
		}
		if strings.EqualFold(filepath.Base(path), "manifest.json") {
			return nil
		}
		body, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		if err := db.MergeJSON(body); err != nil {
			return fmt.Errorf("%s: %w", path, err)
		}
		return nil
	})
	if err != nil {
		return nil, err
	}
	return db, nil
}

func (db *DB) MergeJSON(body []byte) error {
	var data fileData
	if err := json.Unmarshal(body, &data); err != nil {
		return err
	}
	db.MergeData(data)
	return nil
}

func (db *DB) MergeData(data fileData) {
	if db == nil {
		return
	}
	for _, typ := range data.Types {
		db.addType(typ)
	}
	for _, c := range data.Constants {
		db.Constants[fold(c.Name)] = c
	}
	for progID, typ := range data.ProgIDs {
		key := fold(progID)
		db.ProgIDs[key] = typ
		db.ProgIDNames[key] = progID
	}
	for name, typ := range data.GlobalValues {
		db.GlobalValues[fold(name)] = typ
		db.GlobalNames[fold(name)] = name
	}
}

func (db *DB) addType(typ TypeInfo) {
	if typ.Name == "" {
		return
	}
	key := fold(typ.Name)
	if existing, ok := db.Types[key]; ok {
		typ = mergeType(existing, typ)
	}
	db.Aliases[fold(typ.Name)] = typ.Name
	short := typ.Name
	if idx := strings.LastIndex(short, "."); idx >= 0 {
		short = short[idx+1:]
	}
	db.Aliases[fold(short)] = typ.Name
	for _, alias := range typ.Aliases {
		db.Aliases[fold(alias)] = typ.Name
	}
	if typ.Kind == "collection" {
		typ.Collection = true
	}
	db.Types[key] = typ
}

func mergeType(base, overlay TypeInfo) TypeInfo {
	out := base
	if overlay.Name != "" {
		out.Name = overlay.Name
	}
	if overlay.Library != "" {
		out.Library = overlay.Library
	}
	if overlay.Kind != "" {
		out.Kind = overlay.Kind
	}
	if overlay.Summary != "" {
		out.Summary = overlay.Summary
	}
	if overlay.ElementType != "" {
		out.ElementType = overlay.ElementType
	}
	if overlay.DefaultMember != "" {
		out.DefaultMember = overlay.DefaultMember
	}
	if overlay.DefaultMemberType != "" {
		out.DefaultMemberType = overlay.DefaultMemberType
	}
	if overlay.Collection {
		out.Collection = true
	}
	if overlay.Confidence != "" && !preserveGeneratedProvenance(base, overlay) {
		out.Confidence = overlay.Confidence
	}
	if overlay.Source != "" && !preserveGeneratedProvenance(base, overlay) {
		out.Source = overlay.Source
	}
	out.Aliases = mergeStrings(out.Aliases, overlay.Aliases)
	out.AssignableTo = mergeStrings(out.AssignableTo, overlay.AssignableTo)
	out.Properties = mergeMembers(out.Properties, overlay.Properties)
	out.Methods = mergeMembers(out.Methods, overlay.Methods)
	out.Events = mergeMembers(out.Events, overlay.Events)
	return out
}

// preserveGeneratedProvenance keeps a generated TypeLib type authoritative
// when the embedded curated DB supplies convenience metadata or member
// corrections. The curated overlay is intentionally loaded after generated
// databases, but it must not make a complete generated member set look partial.
func preserveGeneratedProvenance(base, overlay TypeInfo) bool {
	return strings.EqualFold(base.Source, "typelib") &&
		strings.EqualFold(base.Confidence, "generated") &&
		strings.EqualFold(overlay.Source, "xlflow") &&
		strings.EqualFold(overlay.Confidence, "curated")
}

func mergeStrings(base, overlay []string) []string {
	if len(base) == 0 {
		return append([]string{}, overlay...)
	}
	out := append([]string{}, base...)
	seen := map[string]bool{}
	for _, value := range out {
		seen[fold(value)] = true
	}
	for _, value := range overlay {
		if value == "" || seen[fold(value)] {
			continue
		}
		seen[fold(value)] = true
		out = append(out, value)
	}
	return out
}

func mergeMembers(base, overlay []MemberInfo) []MemberInfo {
	if len(base) == 0 {
		return append([]MemberInfo{}, overlay...)
	}
	out := append([]MemberInfo{}, base...)
	index := map[string]int{}
	for i, member := range out {
		index[fold(member.Name)] = i
	}
	for _, member := range overlay {
		key := fold(member.Name)
		if key == "" {
			continue
		}
		if i, ok := index[key]; ok {
			out[i] = member
			continue
		}
		index[key] = len(out)
		out = append(out, member)
	}
	return out
}

func (db *DB) ResolveType(name string) (TypeInfo, bool) {
	if db == nil {
		return TypeInfo{}, false
	}
	key := fold(name)
	if canonical, ok := db.Aliases[key]; ok {
		key = fold(canonical)
	}
	typ, ok := db.Types[key]
	return typ, ok
}

func (db *DB) ResolveConstant(name string) (ConstantInfo, bool) {
	if db == nil {
		return ConstantInfo{}, false
	}
	c, ok := db.Constants[fold(name)]
	return c, ok
}

func (db *DB) ResolveProgID(progID string) (TypeInfo, bool) {
	if db == nil {
		return TypeInfo{}, false
	}
	typ, ok := db.ProgIDs[fold(strings.Trim(progID, `"`))]
	if !ok {
		return TypeInfo{}, false
	}
	return db.ResolveType(typ)
}

func (db *DB) ResolveGlobal(name string) (TypeInfo, bool) {
	if db == nil {
		return TypeInfo{}, false
	}
	typ, ok := db.GlobalValues[fold(name)]
	if !ok {
		return TypeInfo{}, false
	}
	return db.ResolveType(typ)
}

func (db *DB) ResolveMember(receiverType, member string) (MemberInfo, bool) {
	typ, ok := db.ResolveType(receiverType)
	if !ok {
		return MemberInfo{}, false
	}
	for _, m := range typ.Properties {
		if strings.EqualFold(m.Name, member) {
			return m, true
		}
	}
	for _, m := range typ.Methods {
		if strings.EqualFold(m.Name, member) {
			return m, true
		}
	}
	for _, m := range typ.Events {
		if strings.EqualFold(m.Name, member) {
			return m, true
		}
	}
	if typ.DefaultMember != "" && strings.EqualFold(member, typ.DefaultMember) {
		return MemberInfo{Name: typ.DefaultMember, ReturnType: typ.ElementType, Default: true}, true
	}
	return MemberInfo{}, false
}

func (db *DB) Members(receiverType string) []MemberInfo {
	typ, ok := db.ResolveType(receiverType)
	if !ok {
		return nil
	}
	seen := map[string]bool{}
	var out []MemberInfo
	add := func(m MemberInfo) {
		if m.Name == "" {
			return
		}
		key := fold(m.Name)
		if seen[key] {
			return
		}
		seen[key] = true
		out = append(out, m)
	}
	for _, m := range typ.Properties {
		add(m)
	}
	for _, m := range typ.Methods {
		add(m)
	}
	for _, m := range typ.Events {
		add(m)
	}
	if typ.DefaultMember != "" {
		add(MemberInfo{Name: typ.DefaultMember, ReturnType: typ.ElementType, Default: true})
	}
	sort.SliceStable(out, func(i, j int) bool { return out[i].Name < out[j].Name })
	return out
}

func (db *DB) IsAssignable(targetType, sourceType string) (assignable bool, known bool) {
	target, ok := db.ResolveType(targetType)
	if !ok {
		return false, false
	}
	source, ok := db.ResolveType(sourceType)
	if !ok {
		return false, false
	}
	if strings.EqualFold(target.Name, source.Name) {
		return true, true
	}
	if len(source.AssignableTo) == 0 {
		return false, false
	}
	seen := map[string]bool{}
	var visit func(TypeInfo) bool
	visit = func(typ TypeInfo) bool {
		for _, parentName := range typ.AssignableTo {
			parent, ok := db.ResolveType(parentName)
			if !ok {
				continue
			}
			key := fold(parent.Name)
			if seen[key] {
				continue
			}
			seen[key] = true
			if strings.EqualFold(parent.Name, target.Name) || visit(parent) {
				return true
			}
		}
		return false
	}
	return visit(source), true
}

func (db *DB) TypeNames() []string {
	if db == nil {
		return nil
	}
	names := make([]string, 0, len(db.Types))
	for _, typ := range db.Types {
		names = append(names, typ.Name)
	}
	sort.Strings(names)
	return names
}

func (db *DB) ConstantsList() []ConstantInfo {
	if db == nil {
		return nil
	}
	out := make([]ConstantInfo, 0, len(db.Constants))
	for _, constant := range db.Constants {
		out = append(out, constant)
	}
	sort.SliceStable(out, func(i, j int) bool { return out[i].Name < out[j].Name })
	return out
}

func (db *DB) ProgIDsList() []string {
	if db == nil {
		return nil
	}
	out := make([]string, 0, len(db.ProgIDs))
	for progID := range db.ProgIDs {
		name := db.ProgIDNames[progID]
		if name == "" {
			name = progID
		}
		out = append(out, name)
	}
	sort.Strings(out)
	return out
}

func (db *DB) GlobalsList() []MemberInfo {
	if db == nil {
		return nil
	}
	out := make([]MemberInfo, 0, len(db.GlobalValues))
	for key, typ := range db.GlobalValues {
		name := db.GlobalNames[key]
		if name == "" {
			name = key
		}
		out = append(out, MemberInfo{Name: name, ReturnType: typ})
	}
	sort.SliceStable(out, func(i, j int) bool { return out[i].Name < out[j].Name })
	return out
}

func fold(s string) string {
	return strings.ToLower(strings.TrimSpace(s))
}
