package vbadb

import (
	"embed"
	"encoding/json"
	"fmt"
	"io/fs"
	"path/filepath"
	"sort"
	"strings"
)

//go:embed builtin/*.json
var builtinFS embed.FS

type DB struct {
	Types        map[string]TypeInfo
	Aliases      map[string]string
	Constants    map[string]ConstantInfo
	ProgIDs      map[string]string
	GlobalValues map[string]string
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
	Name     string `json:"name"`
	Type     string `json:"type,omitempty"`
	Optional bool   `json:"optional,omitempty"`
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
	db := &DB{
		Types:        map[string]TypeInfo{},
		Aliases:      map[string]string{},
		Constants:    map[string]ConstantInfo{},
		ProgIDs:      map[string]string{},
		GlobalValues: map[string]string{},
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
		var data fileData
		if err := json.Unmarshal(body, &data); err != nil {
			return fmt.Errorf("%s: %w", path, err)
		}
		for _, typ := range data.Types {
			db.addType(typ)
		}
		for _, c := range data.Constants {
			db.Constants[fold(c.Name)] = c
		}
		for progID, typ := range data.ProgIDs {
			db.ProgIDs[fold(progID)] = typ
		}
		for name, typ := range data.GlobalValues {
			db.GlobalValues[fold(name)] = typ
		}
		return nil
	})
	if err != nil {
		return nil, err
	}
	return db, nil
}

func (db *DB) addType(typ TypeInfo) {
	if typ.Name == "" {
		return
	}
	db.Types[fold(typ.Name)] = typ
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
	for _, m := range append(append([]MemberInfo{}, typ.Properties...), typ.Methods...) {
		if strings.EqualFold(m.Name, member) {
			return m, true
		}
	}
	if typ.DefaultMember != "" && strings.EqualFold(member, typ.DefaultMember) {
		return MemberInfo{Name: typ.DefaultMember, ReturnType: typ.ElementType, Default: true}, true
	}
	return MemberInfo{}, false
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

func fold(s string) string {
	return strings.ToLower(strings.TrimSpace(s))
}
