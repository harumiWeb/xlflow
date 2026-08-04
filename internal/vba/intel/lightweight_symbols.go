package intel

import "strings"

// LightweightDocumentSymbols resolves exact procedure queries without forcing
// full-document symbol extraction. The boolean reports whether the query shape
// and source catalog were handled; callers should fall back when it is false.
func (a Analyzer) LightweightDocumentSymbols(doc Document, query WorkspaceSymbolQuery) ([]Symbol, bool) {
	if query.Mode != WorkspaceSymbolQueryExact && query.Mode != WorkspaceSymbolQueryQualified {
		return nil, false
	}
	needle := strings.TrimSpace(query.Text)
	if query.Mode == WorkspaceSymbolQueryQualified {
		if index := strings.LastIndex(needle, "."); index >= 0 {
			needle = needle[index+1:]
		}
	}
	if needle == "" {
		return nil, false
	}
	catalog := procedureCatalogForDocumentMode(doc, false)
	if !catalog.ReuseSafe {
		return nil, false
	}
	module := moduleNameForDocument(doc)
	var out []Symbol
	for _, entry := range catalog.Entries {
		if !strings.EqualFold(entry.Identity.CanonicalName, needle) {
			continue
		}
		fragment := doc
		fragment.Source = doc.Source[entry.StartByte:entry.EndByte]
		fragment.Snapshot = nil
		symbols, err := a.DocumentSymbols(fragment)
		if err != nil {
			return nil, false
		}
		for _, symbol := range symbols {
			symbol.Range = rebaseRange(symbol.Range, Position{}, entry.Range.Start)
			symbol.Selection = rebaseRange(symbol.Selection, Position{}, entry.Range.Start)
			symbol.File = doc.Path
			symbol.Module = module
			if strings.EqualFold(symbol.Name, needle) {
				out = append(out, symbol)
			}
		}
	}
	return out, len(out) > 0
}

func rebaseRange(source Range, oldStart, newStart Position) Range {
	return Range{Start: rebasePosition(source.Start, oldStart, newStart), End: rebasePosition(source.End, oldStart, newStart)}
}
