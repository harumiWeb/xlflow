package calls

import (
	"cmp"
	"slices"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

// FromResolvedIR projects already-resolved calls without querying the resolver
// again. Unlike the legacy inspect adapter, it retains canonical resolution
// statuses, including non_callable, incomplete, and dynamic. Lexical shadowing
// and RaiseEvent decisions therefore remain those of the resolved IR snapshot.
func FromResolvedIR(ir procedureir.DocumentIR) []Call {
	parse := symbols.ParseSummary{HasError: ir.Parse.HasError, HasMissing: ir.Parse.HasMissing}
	out := make([]Call, 0)
	for _, procedure := range ir.Procedures {
		for _, site := range procedure.Calls {
			resolution := legacyResolution(site.Resolution)
			resolution.Status = string(site.Resolution.Status)
			out = append(out, Call{CallSite: callSiteFromIR(site, parse), Resolution: resolution})
		}
	}
	slices.SortStableFunc(out, func(a, b Call) int { return cmp.Compare(a.Range.StartByte, b.Range.StartByte) })
	return out
}

// TypeReferencesFromIR projects type references without constructing unused
// call-site projections for consumers that already have resolved calls.
func TypeReferencesFromIR(ir procedureir.DocumentIR) []TypeReference {
	parse := symbols.ParseSummary{HasError: ir.Parse.HasError, HasMissing: ir.Parse.HasMissing}
	out := make([]TypeReference, 0, len(ir.TypeReferences))
	for _, reference := range ir.TypeReferences {
		out = append(out, typeReferenceFromIR(reference, parse))
	}
	slices.SortStableFunc(out, func(a, b TypeReference) int { return cmp.Compare(a.Range.StartByte, b.Range.StartByte) })
	return out
}
