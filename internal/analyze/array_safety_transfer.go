package analyze

import (
	"strconv"
	"strings"
)

func (a Analyzer) arrayTransfer(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard) (arrayFlowState, []Finding) {
	if state == nil {
		state = arrayFlowState{}
	}
	if assignment, ok := inlineArrayAssignmentText(text); ok {
		text = assignment
	}
	var findings []Finding
	addWithKey := func(operationKey, code, message, reason, suggestion string) {
		if code == "VBA227" && !a.Config.Analyze.DetectArrayLifecycleSafety {
			return
		}
		if code == "VBA208" && !a.Config.Analyze.DetectRedimPreserveDimension {
			return
		}
		if code == "VBA209" && !a.Config.Analyze.DetectObjectArrayComparison {
			return
		}
		finding := a.simpleFinding(file, proc, line, code, "warning", message, reason, suggestion)
		finding.arrayLifecycleFinding = code == "VBA227"
		finding.arrayOperationKey = operationKey
		findings = append(findings, finding)
	}
	add := func(code, message, reason, suggestion string) {
		addWithKey("", code, message, reason, suggestion)
	}

	lower := strings.ToLower(strings.TrimSpace(text))
	if declRe.MatchString(text) || isProcedureHeaderLine(lower) {
		return state, findings
	}
	allocationProbeParameter := ""
	if ctx.arrayAllocationGuards[strings.ToLower(proc.Name)] {
		allocationProbeParameter, _ = arrayAllocationGuardParameter(proc)
	}
	if allocationProbeParameter == "" && ctx.arraySafeBoundGuards[strings.ToLower(proc.Name)] {
		allocationProbeParameter, _ = arraySafeBoundGuardParameter(proc)
	}
	if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
		base := arrayOptionBase(file)
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if !direct {
				// Keep the existing receiver-array state transition for nested
				// member ReDim statements. VBA208 must not attribute the member
				// shape to that receiver, but changing the shared VBA227 state is
				// a separate analyzer contract.
				legacy := arrayRedimClauseRe.FindStringSubmatch(clause)
				if len(legacy) == 0 {
					continue
				}
				redim = directArrayRedimClause{name: legacy[1], dimensions: legacy[2]}
			}
			name := strings.ToLower(redim.name)
			variable, known := variables[name]
			// An unresolved target may be an implicit Variant or an external
			// member.  ReDim is valid for some of those runtime shapes, so only
			// report a target mismatch once the shared declaration facts prove a
			// scalar or fixed array.
			if !known {
				continue
			}
			old := state[name]
			dimensions := parseArrayDimensionsWithConstants(redim.dimensions, base, constants)
			// Ordinary analysis keeps an unknown Variant conservative. During
			// documented array-return inference, VBE-confirmed non-Preserve
			// ReDim semantics allow the Variant to establish its array value.
			preserve := strings.TrimSpace(match[1]) != ""
			variantArray := variable.isVariant && (old.knownArray || ctx.arrayAllowVariantRedim && !preserve)
			resizable := (variable.isArray || variantArray) && !variable.fixed
			if variable.isVariant && !old.knownArray && (!ctx.arrayAllowVariantRedim || preserve) {
				continue
			}
			if !resizable {
				// An unresolved object/UDT/external declaration is not a proven
				// scalar.  Leave it unknown rather than guessing that ReDim is
				// invalid; the shared shape contract is deliberately fail-open.
				if !variable.isArray && !variable.isVariant && !variable.knownScalar {
					continue
				}
				add("VBA227", redim.name+" is not a dynamic array and cannot be resized with ReDim.", "ReDim requires a dynamic array; fixed-size arrays and scalar values have no resizable allocation state.", "Declare the value as a dynamic array, or remove ReDim and use its declared bounds.")
			} else if impossibleArrayBounds(dimensions) {
				add("VBA227", redim.name+" has impossible constant ReDim bounds.", "A ReDim lower bound cannot be greater than its upper bound.", "Use bounds whose lower value is less than or equal to the upper value, or keep the bounds dynamic.")
			} else if direct && match[1] != "" && !preserveDimensionsSafe(old.preserveShape, dimensions) {
				add("VBA208", "ReDim Preserve may change a non-final or unknown array dimension.", "VBA can only preserve an array while changing its final dimension, and that cannot be proven when the prior shape is unknown.", "Only change the final dimension, or copy values into a newly sized array explicitly.")
			}
			if resizable && !impossibleArrayBounds(dimensions) {
				next := arrayValue{kind: arrayAllocated, knownArray: true, dimensions: dimensions, origin: arrayOriginLocal}
				if direct {
					next.preserveShape = dimensions
				} else {
					next.preserveShape = append([]arrayDimension(nil), old.preserveShape...)
				}
				state[name] = next
			}
		}
		return state, findings
	}
	if match := arrayEraseRe.FindStringSubmatch(text); len(match) > 0 {
		for _, target := range splitArgs(match[1]) {
			name := strings.ToLower(strings.TrimSpace(target))
			if !arrayEraseNameRe.MatchString(name) {
				continue
			}
			if variable, ok := variables[name]; ok {
				if variable.fixed {
					state[name] = arrayValue{
						kind:          arrayAllocated,
						knownArray:    true,
						dimensions:    append([]arrayDimension(nil), variable.dimensions...),
						preserveShape: append([]arrayDimension(nil), variable.dimensions...),
						origin:        arrayOriginLocal,
					}
				} else if variable.isArray {
					state[name] = arrayValue{kind: arrayUnallocated, knownArray: true, origin: arrayOriginLocal}
				} else if variable.isVariant {
					state[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
				} else if variable.knownScalar || variable.isObject {
					add("VBA227", strings.TrimSpace(target)+" is not an array and cannot be erased as an array.", "Erase applies to arrays; a scalar value has no array allocation to clear.", "Erase an array variable or remove the Erase statement for this scalar value.")
				}
			}
		}
		return state, findings
	}
	// Inline `If ... Then ReDim ...` has a branch-specific state that is not
	// represented by one transfer result. Keep it conservative and avoid
	// treating the ReDim bounds themselves as an array access.
	if strings.Contains(lower, "redim ") {
		return state, findings
	}

	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		argument := strings.TrimSpace(bound[2])
		name := strings.ToLower(argument)
		value, ok := state[name]
		variable, known := variables[name]
		if !known {
			if scalarExpressionKnown(argument) {
				addWithKey(arrayBoundOperationKey(bound[1], argument, "scalar"), "VBA227", bound[1]+" cannot be used on a known scalar expression.", "LBound and UBound require an array value; this argument is a statically known scalar.", "Pass an array value to the bound function or remove the bound query.")
			}
			continue
		}
		if !variable.isArray && !variable.isVariant {
			if !variable.knownScalar && !variable.isObject {
				continue
			}
			addWithKey(arrayBoundOperationKey(bound[1], argument, "scalar"), "VBA227", bound[1]+" cannot be used on non-array "+variable.name+".", "LBound and UBound require an array value; this target is a known scalar.", "Pass an array variable to the bound function or remove the bound query.")
			continue
		}
		if !ok || value.origin == arrayOriginRangeValue {
			continue
		}
		// A Variant has no statically proven array nature.  Keep this path
		// fail-open; only a proven array (or a proven scalar handled above) is
		// actionable here.
		if variable.isVariant && !value.knownArray {
			continue
		}
		if value.kind != arrayAllocated || !value.knownArray {
			if arrayResumeNextCapacityProofApplies(capacityGuards, name, line) {
				// A recognized Resume Next capacity probe deliberately catches
				// this bounds failure before its fallback allocation branch.
				continue
			}
			if allocationProbeParameter != "" && strings.EqualFold(argument, allocationProbeParameter) {
				// A recognized allocation probe deliberately catches this
				// bounds failure and returns zero from its recovery label.
				continue
			}
			addWithKey(arrayBoundOperationKey(bound[1], argument, "unallocated"), "VBA227", bound[1]+" is used before "+variable.name+" is proven to be allocated.", "LBound and UBound raise a runtime error for an unallocated dynamic array and are unsafe for an unknown Variant.", "Allocate the array on every path before querying its bounds, or guard the operation explicitly.")
			continue
		}
		dimension := 1
		if argument := strings.TrimSpace(bound[3]); argument != "" {
			parsed, err := strconv.Atoi(argument)
			if err != nil {
				// A variable or expression is an unknown dimension, not an
				// invalid one. No contradiction can be proven statically.
				continue
			}
			dimension = parsed
		}
		if dimension < 1 || len(value.dimensions) > 0 && dimension > len(value.dimensions) {
			addWithKey(arrayBoundOperationKey(bound[1], argument, "bounds"), "VBA227", bound[1]+" uses invalid dimension "+strconv.Itoa(dimension)+" for "+variable.name+".", "The requested dimension is outside the array dimensions known at this point.", "Use a valid dimension number for the array, or avoid assuming a shape that is not statically known.")
		}
	}

	if match := arrayForEachRe.FindStringSubmatch(text); len(match) > 0 {
		if iterableSourceKnownInvalid(match[1], variables, state, ctx) {
			add("VBA227", strings.TrimSpace(match[1])+" is not a collection or array and cannot be used as a For Each source.", "For Each requires an iterable Collection or array value; this source is a known scalar.", "Iterate an array or Collection, or change the source expression to an iterable value.")
		}
	}

	for _, use := range arrayIndexedUsesForSource(text, variables) {
		// An empty subscript pair passes the whole array to a procedure; it is
		// not an element access whose dimension or allocation should be checked
		// at the call site. The callee owns any element-access diagnostics.
		if len(use.args) == 0 {
			continue
		}
		if arrayResumeNextCapacityIndexApplies(capacityGuards, use.name, line) {
			continue
		}
		value := state[strings.ToLower(use.name)]
		if variable, ok := variables[strings.ToLower(use.name)]; ok && variable.isVariant && !value.knownArray {
			continue
		}
		if value.origin == arrayOriginRangeValue {
			continue
		}
		if value.kind != arrayAllocated || !value.knownArray {
			addWithKey(arrayIndexOperationKey(use.name, "unallocated"), "VBA227", use.name+" is indexed before its array allocation is guaranteed.", "An array access can fail after Erase, before ReDim, or on a branch where allocation is not established.", "Allocate the array on every path before indexing it, or guard the access with a proven allocation check.")
			continue
		}
		if value.mayBeEmpty {
			addWithKey(arrayIndexOperationKey(use.name, "empty"), "VBA227", use.name+" is indexed while its Byte array may be empty.", "A zero-length Byte array has valid bounds queries but no element that can be indexed.", "Guard the element access with a positive length or allocate a non-empty Byte array first.")
			continue
		}
		if len(value.dimensions) > 0 && len(use.args) != len(value.dimensions) {
			addWithKey(arrayIndexOperationKey(use.name, "dimension"), "VBA227", use.name+" is indexed with "+strconv.Itoa(len(use.args))+" dimension(s), but its known shape has "+strconv.Itoa(len(value.dimensions))+".", "The number of subscripts must match the array dimensions known to the analyzer.", "Use the correct number of subscripts or revise the declared array shape.")
			continue
		}
		for i, arg := range use.args {
			if i >= len(value.dimensions) {
				break
			}
			if literal, ok := integerLiteral(arg); ok {
				dimension := value.dimensions[i]
				if dimension.lower.known && literal < dimension.lower.value || dimension.upper.known && literal > dimension.upper.value {
					addWithKey(arrayIndexOperationKey(use.name, "bounds"), "VBA227", use.name+" is indexed outside its known bounds.", "The subscript contradicts the lower or upper bound established by the declaration or ReDim.", "Use an index within the declared bounds, or establish the bounds dynamically before access.")
				}
			}
		}
	}

	if match := arrayForBoundRe.FindStringSubmatch(text); len(match) > 0 {
		name := strings.ToLower(match[2])
		if variable, ok := variables[name]; ok {
			value := state[name]
			if value.kind == arrayAllocated && value.knownArray && len(value.dimensions) > 0 && value.dimensions[0].lower.known {
				if start, ok := integerLiteral(match[1]); ok && start < value.dimensions[0].lower.value {
					add("VBA227", "The loop range assumes an inconsistent lower bound for "+variable.name+".", "The loop starts at a value different from the known lower bound of the array.", "Use LBound("+variable.name+") as the loop start.")
				}
			}
		}
	}

	if lhs, rhs, indexed, ok := arrayAssignment(text); ok {
		name := strings.ToLower(lhs)
		if variable, exists := variables[name]; exists && !variable.isArray && !variable.isVariant {
			if argument, probe := arrayAllocationProbeArgument(rhs, ctx.arrayAllocationGuards); probe {
				state[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginLocal, allocationProbe: argument}
			} else if argument, probe := arraySafeBoundProbeArgument(rhs, ctx.arraySafeBoundGuards); probe {
				state[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginLocal, safeBoundProbe: argument}
			} else if value, tracked := state[name]; tracked && (value.allocationProbe != "" || value.safeBoundProbe != "") {
				value.allocationProbe = ""
				value.safeBoundProbe = ""
				state[name] = value
			}
		}
		if variable, exists := variables[name]; exists && variable.isArray && variable.isObject && indexed && !strings.HasPrefix(lower, "set ") {
			code := "VBA101"
			typ := variable.typ
			callee := arrayCallName(rhs)
			if returnType := ctx.functionReturns[callee]; isObjectType(returnType) {
				code = "VBA102"
				typ = returnType
			}
			if code == "VBA101" || code == "VBA102" {
				// These are existing missing-Set rules; the lifecycle rule only
				// supplies the indexed target that the old text matcher missed.
				findings = append(findings, a.objectSetFinding(file, proc, line, code, strings.TrimSpace(lhs), typ))
			}
		}
		if !indexed {
			if value, known := arrayDictionaryMemberExpressionState(file, proc, line, rhs, variables); known {
				if variable, exists := variables[name]; exists && (variable.isArray || variable.isVariant) {
					state[name] = value
				}
			} else if value, known := arrayExpressionState(rhs, state, ctx); known {
				if value.mayBeEmpty && arrayExpressionKnownNonEmpty(file, proc, line, rhs, variables) {
					value.mayBeEmpty = false
				}
				if variable, exists := variables[name]; exists && (variable.isArray || variable.isVariant) {
					state[name] = value
				}
			} else if variable, exists := variables[name]; exists {
				if value, assigned := byteArrayStringAssignment(file, proc, line, variable, rhs, variables); assigned {
					state[name] = value
				} else if variable.isArray || variable.isVariant {
					value := arrayValue{kind: arrayUnknown, knownArray: false, origin: arrayOriginUnknown}
					// An arbitrary source assigned to a Byte array may be a
					// zero-length byte array. Keep that possibility so a later
					// successful bounds query can prove allocation without
					// incorrectly proving that an element exists.
					if isByteArrayVariable(variable) {
						value.mayBeEmpty = true
					}
					state[name] = value
				}
			}
		}
	}
	return state, findings
}
