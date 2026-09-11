using Danfma.MySheet.Expressions.Logical;
using Danfma.MySheet.Expressions.Lookup;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// The rectangular result of evaluating a sub-expression element-by-element: a vector of
/// <see cref="ComputedValue"/> laid out ROW-MAJOR with its <see cref="Rows"/>×<see cref="Columns"/>
/// dimensions. Errors and logicals are preserved per element — the CONSUMER decides how to fold/propagate
/// them (Phase B); this struct never collapses the vector to a scalar.
/// </summary>
internal readonly struct ArrayEvaluationResult
{
    public int Rows { get; }
    public int Columns { get; }
    public ComputedValue[] Values { get; }

    public ArrayEvaluationResult(int rows, int columns, ComputedValue[] values)
    {
        Rows = rows;
        Columns = columns;
        Values = values;
    }

    /// <summary>The element count (<c>Rows * Columns</c>, i.e. <c>Values.Length</c>).</summary>
    public int Length => Values.Length;
}

/// <summary>
/// The internal "mini-CSE" element-wise evaluator (Phase A of <c>plans/mini-cse-array-arguments.md</c>).
/// Given an AST node and a context it reproduces Excel's implicit array/CSE semantics WITHOUT any public
/// array value, spilling or new AST node: a closed range becomes a vector of its per-cell values, scalars
/// broadcast, <c>BinaryOperation</c> zips element-wise, <c>IF</c> zips a condition array against its
/// branches (a branch-less <c>IF</c> yields a logical <c>FALSE</c> where the condition is false — the idiom
/// <c>SMALL(IF(...))</c> depends on this), and <c>ROW</c>/<c>COLUMN</c> of a rectangle becomes a vector of
/// row/column numbers — over a range written literally, and over a defined name or any other node that
/// DENOTES one (a structured reference, a <c>':'</c> range with reference-returning endpoints), whose shape
/// is discovered by resolving it. A bare defined NAME in an array position is whatever it is bound to
/// (Phase 11a Rule A, <see cref="ResolveNameShape"/>): a rectangle streams its cells exactly like the
/// literal — a rectangle on a MISSING sheet included, so the per-element <c>#REF!</c> streams — an open
/// range is refused, a single cell or a union broadcasts the resolved node's value, and a constant, a
/// formula name that does NOT resolve to a reference, or an unknown name broadcasts the name's own value
/// (<c>#NAME?</c> included) — a formula name that DOES resolve to one is that reference, and takes the
/// rectangle/open-range/single-cell arm accordingly. At a
/// consumer's TOP level a bare name is a reference and takes the reference path exactly as a bare literal
/// range does: <see cref="IsBareReferenceNode(Expression, EvaluationContext)"/> is the one predicate the
/// three top-level gates share, and DefinedNameArrayEligibilityTests' must-not-move pins are what make it
/// load-bearing — unless the name is LET-bound to an ARRAY (Phase 11c, <see cref="ArrayBindings"/>), which
/// the predicate does not admit: the same <c>f</c> then streams to <c>SUM(f)</c> and <c>ROWS(f)</c> and is
/// refused by <c>COUNTIF(f,…)</c>, both right (ArrayBindingTests). Two shapes are LIFTED
/// element-wise (Phase 8): a unary <c>-</c>/<c>%</c>
/// over an array (unary <c>+</c> is Excel's reference-preserving no-op and stays opaque, so
/// <c>SUM(+A1:A3)</c> keeps reading the cells), and any built-in the registry classifies
/// <see cref="ArrayLifting.Elementwise"/> — a pure-scalar function such as <c>LEN</c>, <c>ROUND</c> or
/// <c>IFERROR</c> — with at least one array argument, whose scalar body is evaluated once per element over
/// rebound argument slots (<see cref="LiftedFunctionOperand"/>) while its scalar arguments broadcast. Any
/// node that PRODUCES an array — Phase 7's <c>FILTER</c>/<c>SORT</c>/<c>UNIQUE</c>/<c>SEQUENCE</c> — plugs in
/// through <see cref="IArrayProducer"/>, reached by one arm in each of <see cref="Probe"/> and
/// <see cref="TryBuildOperand"/> and answering for its own shape and elements (a bare producer in a cell
/// is its top-left element, <see cref="FirstElement(Expression, EvaluationContext)"/>). A <c>LET</c> is an
/// array when its body is one in the scope its bindings make (<see cref="ArrayBindings"/>; the <c>Let</c>
/// arms of <see cref="Probe"/> and <see cref="TryBuildOperand"/>). Any
/// node outside this set is treated as a scalar (broadcast); a whole-column/open range is REFUSED (the cost
/// guard) so the whole evaluation reports "not an array" and the caller keeps its current scalar path — except
/// INSIDE a lifted shape, where the refusal makes that unary/function an opaque scalar (evaluated once, its
/// own <c>#VALUE!</c> broadcast) rather than unwinding an enclosing array expression that tolerated it before.
///
/// <para>Two consumption shapes share ONE recursive builder (<see cref="TryBuildOperand"/>): the LAZY
/// <see cref="ArrayStream"/> (element-on-demand, no vector — used by the aggregating consumers SUM/COUNT/
/// AVERAGE/MIN/MAX, SMALL/LARGE and INDEX so a 50k-row idiom allocates only the handful of tree nodes, not a
/// <c>ComputedValue[50k]</c>), and the eager <see cref="ArrayEvaluationResult"/> (materialized once from the
/// same tree) that the direct unit tests drive. Because both walk the identical operand tree, the streamed
/// element sequence is bit-for-bit the materialized vector, in the SAME row-major order — so error
/// propagation ("first error, scan order"), broadcasting (<see cref="Broadcasting"/>: a vector repeats
/// along its 1-long axis, an uncovered position is <c>#N/A</c>) and the evaluate-scalar-operands-once rule
/// (a volatile branch draws once and broadcasts) are preserved.</para>
/// </summary>
internal static class ArrayEvaluation
{
    /// <summary>
    /// Tries to evaluate <paramref name="expression"/> as an array, MATERIALIZING it into a row-major vector.
    /// Returns <c>true</c> (with the vector in <paramref name="result"/>) only when the sub-tree genuinely
    /// produces an array — a closed range, a <c>ROW</c>/<c>COLUMN</c> of one, or an operation/IF with at least
    /// one array operand. A bare scalar, a non-eligible node, or anything touching an open range returns
    /// <c>false</c> (the current scalar path is untouched). Production consumers use
    /// <see cref="TryEvaluateStream"/>; this
    /// eager form exists for the direct unit tests of the vector semantics.
    /// </summary>
    public static bool TryEvaluate(
        Expression expression,
        EvaluationContext context,
        out ArrayEvaluationResult result
    )
    {
        if (TryBuildOperand(expression, context, out var operand) && operand.IsArray)
        {
            var rows = operand.Rows;
            var columns = operand.Columns;
            var values = new ComputedValue[rows * columns];

            for (var index = 0; index < values.Length; index++)
            {
                values[index] = operand.At(index, rows, columns);
            }

            result = new ArrayEvaluationResult(rows, columns, values);
            return true;
        }

        result = default;
        return false;
    }

    /// <summary>
    /// Tries to build a LAZY element-wise view of <paramref name="expression"/> — the allocation-free twin of
    /// <see cref="TryEvaluate"/>. Succeeds under the exact same conditions, but instead of a materialized
    /// vector it returns an <see cref="ArrayStream"/> that computes each element on demand from a small
    /// operand tree (resolved once): the aggregating consumers enumerate it without ever allocating a
    /// <c>ComputedValue[]</c>. Scalar sub-expressions are still evaluated ONCE at build time (broadcast), so a
    /// volatile branch/operand draws a single value and its taint lands in the enclosing cell frame exactly as
    /// before.
    /// </summary>
    public static bool TryEvaluateStream(
        Expression expression,
        EvaluationContext context,
        out ArrayStream stream
    )
    {
        if (TryBuildOperand(expression, context, out var operand) && operand.IsArray)
        {
            stream = new ArrayStream(operand, operand.Rows, operand.Columns);
            return true;
        }

        stream = default;
        return false;
    }

    /// <summary>
    /// A CHEAP pre-check — it never evaluates <paramref name="expression"/> itself — for whether it would
    /// produce an array through the eligible structural set. Consumers gate on this BEFORE calling
    /// <see cref="TryEvaluate"/>/<see cref="TryEvaluateStream"/>, so the scalar hot path pays only a shallow
    /// type-walk and never a double evaluation: when this returns <c>false</c> the consumer keeps its existing
    /// scalar path untouched; when it returns <c>true</c> the subsequent build is guaranteed to succeed and is
    /// the SINGLE evaluation of the argument. It mirrors <see cref="Probe"/>/<see cref="TryBuildOperand"/>
    /// exactly (same array-producing cases, same open-range refusal), so it is true iff the build succeeds as
    /// an array. The array-producing cases are a closed range, a defined name bound to one (a missing
    /// sheet included), <c>ROW</c>/<c>COLUMN</c> of either, a <c>BinaryOperation</c>/<c>IF</c>/unary
    /// <c>-</c>/<c>%</c> with an array operand, and an <see cref="ArrayLifting.Elementwise"/> built-in with an
    /// array argument (the lift inspects only the registry classification and recurses into the ARGUMENTS'
    /// eligibility — it evaluates nothing, measured with a counting custom function in the argument slots).
    /// </summary>
    /// <remarks>
    /// <para>It takes a <paramref name="context"/> because the check is no longer purely SYNTACTIC: whether
    /// <c>ROW</c>/<c>COLUMN</c> of a defined name (or of any other node that merely DENOTES a reference) is an
    /// array or a scalar depends on the SHAPE the argument resolves to — a rectangle vs a single cell — which
    /// no type-walk can know, and the same is true of a bare name itself. Those cases cost one reference
    /// RESOLUTION — a dictionary lookup plus a virtual call for a name, never an evaluation of the
    /// <c>ROW</c>/<c>COLUMN</c> node itself, though resolving a <c>':'</c> range with reference-returning
    /// endpoints does evaluate those endpoints' own arguments — and they are the only cases that cost more
    /// than the type-walk. That cost normally buys something, because a <c>true</c> answer is followed by the
    /// build that reuses the shape; the exception is <c>Index.TryResolveReference</c>, which probes only to
    /// REJECT the array forms and never builds, so there the resolution (a <c>':'</c> range's endpoint
    /// arguments included) is spent and thrown away. See <see cref="ResolvePositionRange"/> and
    /// <see cref="ResolveNameShape"/>.</para>
    ///
    /// <para>Because a range-bound name IS array-eligible, a consumer gate must exclude a bare
    /// <see cref="NameReference"/> at its top level exactly as it excludes a bare <see cref="Reference"/>
    /// — through <see cref="IsBareReferenceNode(Expression, EvaluationContext)"/>, never a hand-written
    /// <c>is not Reference</c>, which a name (an <see cref="Expression"/>, not a <see cref="Reference"/>)
    /// slips past. Measured on the
    /// prototype WITHOUT that exclusion, fifteen top-level shapes regressed (<c>SUBTOTAL(9,Rng)</c> 14 →
    /// <c>#VALUE!</c>, <c>AGGREGATE(14,0,Nested,1)</c> 2 → a silent 3, <c>SUM(A1:INDEX(Rng,3))</c> 14 →
    /// <c>#REF!</c>, <c>SUM(Wide)</c> on an error fixture <c>#DIV/0!</c> → <c>#N/A</c>, …); they are pinned
    /// in DefinedNameArrayEligibilityTests.</para>
    /// </remarks>
    public static bool IsArrayEligible(Expression expression, EvaluationContext context) =>
        Probe(expression, context).IsArray;

    /// <summary>
    /// Excel's <c>@</c> rule for an ARRAY: the value a cell shows for an array expression is its TOP-LEFT
    /// element. This is the single <see cref="Expression.Evaluate"/> body of every <see cref="IArrayProducer"/>
    /// record — <c>=FILTER(A1:B3,A1:A3>0)</c> in a cell is 5, <c>=SORT(A1:B3,1,-1)</c> is 9,
    /// <c>=SEQUENCE(2,3,7,1)</c> is 7 and <c>=SEQUENCE(2,3)*10</c> is 10 (Aspose.Cells 26.6.0, 2026-09-10,
    /// both entry modes; the contract pins are in <c>ArrayProducerContractTests</c>) — so the rule cannot
    /// drift between the four. It builds the operand (the node's single evaluation, exactly as a consumer's
    /// stream would be) and reads position 0 at the operand's own extent; when the build is refused (an open
    /// range below — the cost guard) or does not yield an array there is no element to take, and the answer
    /// is <c>#VALUE!</c>. It is a rule for PRODUCERS, not for ranges: a bare range in a cell intersects
    /// (<c>ImplicitIntersection</c>), and a range operand under an operator keeps today's <c>#VALUE!</c>
    /// (<c>CellBoundaryIntersectionTests</c>).
    /// </summary>
    internal static ComputedValue FirstElement(Expression expression, EvaluationContext context) =>
        TryBuildOperand(expression, context, out var operand) && operand.IsArray
            ? FirstElement(operand)
            : ComputedValue.Error(Error.Value);

    /// <summary>
    /// The same <c>@</c> rule over an operand that was ALREADY built — an array binding read bare
    /// (<see cref="NameReference.Evaluate"/>, <see cref="ArrayBindings.Binding.TopLeft"/>): position 0 at
    /// the operand's own extent, with no second build.
    /// </summary>
    internal static ComputedValue FirstElement(ArrayOperand operand) =>
        operand.At(0, operand.Rows, operand.Columns);

    /// <summary>
    /// THE mini-CSE consumer gate: the three conditions every consumer that wants to STREAM a computed array
    /// applies, in the one order that is correct. Succeeds (with the lazy view in <paramref name="stream"/>)
    /// only for a non-reference argument that genuinely produces an array; otherwise the caller keeps its own
    /// existing path untouched.
    ///
    /// <para>It lives here as one method because the three conditions had drifted apart once already: a
    /// second hand-written copy is how a consumer silently loses a condition, and the copies cannot be
    /// diffed when they sit in five files.</para>
    ///
    /// <para>The ORDER is load-bearing, not stylistic:</para>
    /// <list type="number">
    /// <item><description><see cref="IsBareReferenceNode(Expression, EvaluationContext)"/> FIRST. A plain
    /// <see cref="RangeReference"/> — an <see cref="AnchoredRangeReference"/>, which <see cref="Probe"/>
    /// classifies as <c>(true, true)</c> exactly like one, and a <see cref="NameReference"/> bound to a
    /// rectangle — IS array-eligible, so without this guard (or with it placed after the probe) every
    /// reference argument would be diverted off its reference path into the stream, losing whatever that
    /// path carries: the snapshot/dense walk, the engine's column-major first-error scan, and for the
    /// aggregate family the nested-SUBTOTAL/AGGREGATE skip, which only the cell-by-cell scan can apply. The
    /// name half of that is measured, not inferred: see <see cref="IsArrayEligible"/>'s remarks. The
    /// predicate is context-aware for exactly one reason: a name LET-bound to an ARRAY is not a reference,
    /// and must reach the stream (Phase 11c).</description></item>
    /// <item><description><see cref="IsArrayEligible"/> SECOND. It is the CHEAP structural pre-check that
    /// never evaluates the expression, so a scalar argument pays only a shallow type-walk before falling
    /// through to the caller's scalar path.</description></item>
    /// <item><description><see cref="TryEvaluateStream"/> LAST, and only once the probe said yes — so it is
    /// the argument's SINGLE evaluation. A volatile operand therefore draws exactly once.</description></item>
    /// </list>
    ///
    /// <para>One nearby site deliberately does NOT use this gate: <c>Index.TryResolveReference</c> probes only
    /// to REJECT the array forms and never builds a stream at all, so it applies the first two conditions
    /// itself — with the same <see cref="IsBareReferenceNode(Expression, EvaluationContext)"/> predicate.
    /// <c>NumericAggregation.Fold</c>'s
    /// <c>default:</c> arm, whose own switch owns the dispatch of every reference NODE, routes through this
    /// gate too: its leading condition is what keeps a bare name (which that switch does not peel off — it is
    /// not a <see cref="Reference"/>) on the referenced-cell path.</para>
    /// </summary>
    public static bool TryStream(
        Expression expression,
        EvaluationContext context,
        out ArrayStream stream
    )
    {
        if (!IsBareReferenceNode(expression, context) && IsArrayEligible(expression, context))
        {
            return TryEvaluateStream(expression, context, out stream);
        }

        stream = default;
        return false;
    }

    /// <summary>
    /// Whether <paramref name="expression"/> is a bare reference NODE — a syntactic <see cref="Reference"/>
    /// or a <see cref="NameReference"/> that is NOT LET-bound to an array in <paramref name="context"/> —
    /// which at a consumer's TOP level must take the consumer's reference path even when
    /// <see cref="IsArrayEligible"/> would say yes for it. The one predicate every top-level gate shares
    /// (<see cref="TryStream"/>, <c>Index.TryResolveReference</c>, <c>NumericAggregation.Fold</c>'s
    /// <c>default:</c> arm, the criteria family's range-slot rejection, the If-branch rule of
    /// <see cref="ProbeIfBranches"/>, and <see cref="ArrayBindings.Capture"/>), so a reference-denoting node
    /// that is not a <see cref="Reference"/> — a name today, a structured table reference tomorrow — is
    /// excluded in ONE place. Measured on the prototype without the name half: fifteen top-level shapes
    /// regressed (see <see cref="IsArrayEligible"/>'s remarks).
    ///
    /// <para>The context is what makes a name's answer honest (Phase 11c): a name whose nearest LET binding
    /// is an ARRAY (<see cref="EvaluationContext.TryGetArrayBinding"/>) is not a reference — it is an array
    /// — and answers <c>false</c>, so <c>SUM(f)</c>/<c>ROWS(f)</c> stream it and <c>COUNTIF(f,…)</c> refuses
    /// it (ArrayBindingTests). A name bound to a scalar or a range, a defined name and an unknown name keep
    /// the syntactic answer, <c>true</c>. There is no context-free overload: every gate runs where a
    /// context exists, and a LET binding does not exist before evaluation, so a caller without one has no
    /// name to ask about.</para>
    /// </summary>
    internal static bool IsBareReferenceNode(Expression expression, EvaluationContext context) =>
        expression switch
        {
            NameReference name => !context.TryGetArrayBinding(name.Name, out _),
            _ => expression is Reference,
        };

    // The pure-shape twin of TryBuildOperand: decides, WITHOUT evaluating the expression, whether the build
    // would succeed (Succeeds — no refused open range on the eligible path) and whether the result is an array
    // (IsArray). Must track the builder's structure exactly so IsArrayEligible == (build result).
    //
    // One case cannot be answered from the syntax alone — ROW/COLUMN over a node that only DENOTES a
    // reference — so both this probe and the builder ask the SAME oracle, ResolvePositionRange, on the same
    // argument in the same context, and therefore agree. The oracle resolves the argument rather than
    // evaluating the ROW/COLUMN node, and the build resolves it a second time, which is what keeps a
    // reference-returning FUNCTION argument out of that arm: see ResolvePositionRange.
    //
    // Internal, not private, so an IArrayProducer in another file can recurse into its own child
    // arguments; the two are the producers' only way in, and IsArrayEligible/TryStream stay the gates.
    internal static (bool Succeeds, bool IsArray) Probe(
        Expression expression,
        EvaluationContext context
    )
    {
        switch (expression)
        {
            case RangeReference:
                return (true, true);

            // Phase 2 audit (shared-formula delta production): a range written INSIDE a shared-formula
            // master is an anchored node, not a plain RangeReference — without this arm it fell through to
            // the `default` case below (an opaque scalar, evaluated once via AnchoredRangeReference.Evaluate,
            // which always yields #VALUE! since a range has no scalar value), silently breaking the mini-CSE
            // idiom (e.g. SMALL(IF(A1:A3=…, …))) for a shared-formula slave.
            case AnchoredRangeReference:
                return (true, true);

            // An open/whole-column range in an array position is refused (the cost guard).
            case OpenRangeReference:
                return (false, false);

            // A name LET-bound to an ARRAY (Phase 11c) is that array: the operand was built once at binding
            // time (ArrayBindings.Capture), so the probe answers for it without resolving anything. Checked
            // BEFORE ResolveNameShape, which would answer Opaque for it (an array binding is not a
            // reference), and before a defined name of the same spelling, which the binding shadows.
            case NameReference bound when context.TryGetArrayBinding(bound.Name, out _):
                return (true, true);

            // A bare defined name is whatever it is bound to (Phase 11a Rule A): the four outcomes of
            // ResolveNameShape map onto the three answers above plus the opaque scalar of `default`. Placed
            // with the reference arms and BEFORE the Row/Column ones (different node types, no shadowing);
            // a name in a ROW/COLUMN argument is that arm's business, not this one's.
            case NameReference:
                return ResolveNameShape(expression, context, out _) switch
                {
                    NameShape.Range => (true, true),
                    NameShape.Refused => (false, false),
                    _ => (true, false),
                };

            // ROW(x)/COLUMN(x) over a reference: ONE arm each, because the two functions differ only in the
            // AXIS the shared shape walk reports — see ProbePosition. The `[NameReference or Reference]`
            // pattern admits every reference NODE (RangeReference, AnchoredRangeReference and
            // OpenRangeReference all derive from Reference) plus a defined name, and deliberately not a
            // reference-returning FUNCTION: see ResolvePositionRange.
            case Row { Arguments: [NameReference or Reference] } row:
                return ProbePosition(row.Arguments[0], context);

            case Column { Arguments: [NameReference or Reference] } column:
                return ProbePosition(column.Arguments[0], context);

            // Unary '-'/'%' is an array exactly when its operand is. Plus is excluded by PATTERN, not by an
            // `if` inside: `+range` must reach `default` and stay the opaque scalar that carries the
            // reference (UnaryOperation.Evaluate routes it through CaptureValue). A refused operand makes
            // the unary an opaque scalar, not a refusal — see ProbeLift.
            case UnaryOperation { Operator: not UnaryOperator.Plus } unary:
                return (true, ProbeLift([unary.Operand], context));

            // A pure-scalar built-in over its arguments — placed AFTER the Row/Column/If arms so those keep
            // their dedicated handling (all three are classified Consumes, so the order is belt-and-braces
            // rather than load-bearing, but a `when` guard silently shadowing If would be a hard bug to find).
            case Function function when TryGetLift(function, out var arguments):
                return (true, ProbeLift(arguments, context));

            case BinaryOperation binary:
            {
                var left = Probe(binary.Left, context);
                if (!left.Succeeds)
                {
                    return (false, false);
                }

                var right = Probe(binary.Right, context);
                if (!right.Succeeds)
                {
                    return (false, false);
                }

                return (true, left.IsArray || right.IsArray);
            }

            // An IF is an array when its CONDITION is one (the element-wise zip) or when a BRANCH is a
            // computed array (a scalar condition then selects that branch whole). The two halves are one
            // arm because both branches are probed either way: a refused branch refuses the IF under
            // either kind of condition. See ProbeIfBranches for why a bare-reference branch does not count.
            case If ifNode when ifNode.Arguments.Length is 2 or 3:
            {
                var condition = Probe(ifNode.Arguments[0], context);
                if (!condition.Succeeds)
                {
                    return (false, false);
                }

                var branches = ProbeIfBranches(ifNode, context, condition.IsArray);
                if (!branches.Succeeds)
                {
                    return (false, false);
                }

                return (true, condition.IsArray || branches.IsArray);
            }

            // A LET is an array when its BODY is one in the scope its bindings make (Phase 11c). Reachable
            // because LET is Entry<Let> — Consumes — so the lift arm above does not take it. See ProbeLet.
            case Let let:
                return ProbeLet(let, context);

            // A node that PRODUCES an array (Phase 7: FILTER/SORT/UNIQUE/SEQUENCE, and any later producer)
            // answers for itself. Placed LAST before `default` and after every arm above: the open-range
            // refusal and the name arm come first because they are different node types; the lift arm
            // comes first because a producer is a Function too, and only its Consumes classification keeps
            // it out of TryGetLift — see IArrayProducer's remarks. Phase 5's TableReference arm and any
            // future reference arm land ABOVE this one.
            case IArrayProducer producer:
                return producer.ProbeArray(context);

            // Anything else is an opaque scalar: succeeds (evaluated once when actually built), not an array.
            default:
                return (true, false);
        }
    }

    // The branch half of the If arm, over the COMPUTED branches only: IsArray when any of them is an array
    // (a producer, a lifted built-in or an operator over a range or a producer), Succeeds when none is
    // refused (the cost guard, as under an array condition). The predicate is exactly TryStream's own —
    // IsBareReferenceNode, then the probe — so the two cannot disagree about what "a computed array" means.
    //
    // A bare reference NODE (a range, a defined name, an open range) is skipped outright: neither counted
    // nor refused. Whether IF(TRUE,A1:A3,0) hands its consumer the RANGE is the "IF returns a reference"
    // question, which the criteria family's gate (CriteriaScan.RejectComputedArray) and the cell-boundary
    // rule both pin at today's answers; keeping every bare-reference branch outside the array path keeps
    // that answer independent of what the OTHER branch holds — SUM(IF(TRUE,A1:A3,0)) and
    // SUM(IF(TRUE,A1:A3,SEQUENCE(3))) are both #VALUE! (the oracle says 14 for both), SUM(IF(TRUE,Rng,…))
    // is 14 through the reference value If.Evaluate carries, and an open range in the UNTAKEN branch
    // costs nothing (SUM(IF(FALSE,A:A,0)*B1:B3) stays 0). TryBuildScalarConditionIf is the other half of
    // that rule: a taken bare-reference branch on the eligible path DECLINES rather than streams.
    //
    // The probe cannot know which branch a scalar condition will TAKE without evaluating it, so it answers
    // for the union: IF(FALSE,SEQUENCE(3),0) is an array here, and the build keeps the promise by handing
    // back the taken scalar as a 1x1 array (the oracle reads it the same way: SUM 0, ROWS 1, COUNTIF #REF!,
    // Aspose.Cells 26.6.0, 2026-09-10, both modes).
    private static (bool Succeeds, bool IsArray) ProbeIfBranches(
        If ifNode,
        EvaluationContext context,
        bool conditionIsArray
    )
    {
        var isArray = false;

        for (var i = 1; i < ifNode.Arguments.Length; i++)
        {
            var branch = ifNode.Arguments[i];

            // A bare reference is never COUNTED toward IsArray — that is the "IF returns a reference"
            // question, which this method declines to answer. Whether it is PROBED depends on the condition,
            // because that is what decides how many branches the build touches:
            //   - ARRAY condition: the build zips, so it builds BOTH branches through TryBuildOperand. A
            //     refusal there must refuse the probe too, or Probe promises an array the build cannot
            //     deliver and the consumer falls back and re-evaluates the condition — a volatile drawn
            //     twice. Found by the phase's final review on IF(RAND()>A1:A3,B:B,0).
            //   - SCALAR condition: the build touches ONE branch, and TryBuildScalarConditionIf answers an
            //     open range with the loud 1x1 #VALUE! WrapScalar gives it rather than refusing. So a
            //     refusable reference in the UNTAKEN branch must cost nothing, which is what the oracle says:
            //     SUM(IF(FALSE,MyColumn,0)*B1:B3) is 0, and probing it here would make it #VALUE!.
            if (IsBareReferenceNode(branch, context) && !conditionIsArray)
            {
                continue;
            }

            var probe = Probe(branch, context);
            if (!probe.Succeeds)
            {
                return (false, false);
            }

            if (IsBareReferenceNode(branch, context))
            {
                continue;
            }

            isArray |= probe.IsArray;
        }

        return (true, isArray);
    }

    // The Let arm's probe half (Phase 11c): walk the bindings with ArrayBindings.Shape — each name bound
    // to the SHAPE its capture would have, an array stand-in or the reference it resolves to or an opaque
    // blank, with nothing evaluated — and probe the body in that scope. A malformed LET is the scalar
    // path's own #VALUE!: an opaque scalar. A bare-reference BODY is not counted, on ProbeIfBranches' rule
    // for a bare-reference branch: whether LET(f,A1:A3,f) hands its consumer the RANGE is the same
    // "returns a reference" question as IF's, left where it is (sweep item 32) — the body still evaluates
    // in the bound scope and the consumer keeps today's answer. The context-aware predicate is what makes
    // that rule and the array case agree about a name: LET(f,FILTER(…),f) has an ARRAY body, and streams.
    //
    // The probe never refuses on a binding: a binding whose expression the cost guard refuses (an open
    // range in an array position) is captured by CaptureValue on the build side, a scalar, and Shape binds
    // it as one — so a refusal is only ever the BODY's, exactly as TryBuildLet finds it.
    private static (bool Succeeds, bool IsArray) ProbeLet(Let let, EvaluationContext context)
    {
        if (!let.TryBind(context, ArrayBindings.Shape, out var scope))
        {
            return (true, false);
        }

        var body = let.Arguments[^1];

        return IsBareReferenceNode(body, scope) ? (true, false) : Probe(body, scope);
    }

    // The Let arm's build half: the SAME walk with ArrayBindings.Capture — the bindings' single evaluation,
    // exactly Let.Evaluate's — then the body built in that scope. A bare-reference body evaluates in the
    // bound scope and broadcasts (the probe did not count it); a malformed LET is the loud #VALUE! scalar,
    // with the well-formed prefix of its bindings evaluated once, as Let.Evaluate evaluates it.
    private static bool TryBuildLet(Let let, EvaluationContext context, out ArrayOperand operand)
    {
        if (!let.TryBind(context, ArrayBindings.Capture, out var scope))
        {
            operand = new ScalarOperand(ComputedValue.Error(Error.Value));
            return true;
        }

        var body = let.Arguments[^1];

        if (IsBareReferenceNode(body, scope))
        {
            operand = new ScalarOperand(body.Evaluate(scope));
            return true;
        }

        return TryBuildOperand(body, scope, out operand);
    }

    // The shape twin of TryBuildPositionOperand — same argument order, same three outcomes, so ROW/COLUMN
    // stay eligible exactly when their build succeeds as an array. A rectangle (written literally, or written
    // inside a shared-formula master as an anchored node) is an array with no resolution at all; an open range
    // is the cost-guard refusal; anything else asks the oracle, whose Array/Refused/Scalar map onto the same
    // three answers the builder gives (a Scalar being the opaque-scalar answer: it succeeds, broadcast).
    private static (bool Succeeds, bool IsArray) ProbePosition(
        Expression argument,
        EvaluationContext context
    ) =>
        argument switch
        {
            RangeReference or AnchoredRangeReference => (true, true),
            OpenRangeReference => (false, false),
            _ => ResolvePositionRange(argument, context, out _) switch
            {
                PositionArgumentShape.Array => (true, true),
                PositionArgumentShape.Refused => (false, false),
                _ => (true, false),
            },
        };

    internal static bool TryBuildOperand(
        Expression expression,
        EvaluationContext context,
        out ArrayOperand operand
    )
    {
        switch (expression)
        {
            case RangeReference range:
                operand = BuildRange(range, context);
                return true;

            // Phase 2 audit (shared-formula delta production): mirrors the RangeReference case above for a
            // range written INSIDE a shared-formula master (an anchored node) — resolved to its concrete,
            // delta-applied twin so BuildRange sees an ordinary RangeReference, no logic duplicated.
            case AnchoredRangeReference anchoredRange:
                operand = BuildRange(anchoredRange.ToRangeReference(context), context);
                return true;

            // Whole-column / whole-row / one-sided open reference: the cost guard keeps it OUT of the
            // mini-CSE. Refuse so the whole evaluation reports "not an array".
            case OpenRangeReference:
                operand = null!;
                return false;

            // The twin of Probe's array-binding arm (Phase 11c): the operand built once at binding time IS
            // the name's operand — no second build, which is what keeps a volatile binding at one draw
            // (LET(x,SEQUENCE(3,1,RAND(),0),SUM(x)-SUM(x)) = 0, ArrayBindingTests).
            case NameReference bound when context.TryGetArrayBinding(bound.Name, out var binding):
                operand = binding;
                return true;

            // The twin of Probe's NameReference arm, on the same oracle: a rectangle streams its cells (a
            // missing sheet included — BuildRange hands back the per-element #REF!); the cost guard refuses
            // an open range; a single cell or a union broadcasts the RESOLVED node's own scalar answer
            // (error included); a constant, a formula name that does NOT resolve to a reference, or an
            // unknown name broadcasts the name's own value, which is where #NAME? still flows through. A
            // formula name that DOES resolve to one takes the arm its resolved reference belongs to.
            case NameReference:
            {
                switch (ResolveNameShape(expression, context, out var resolved))
                {
                    case NameShape.Range:
                        operand = BuildRange((RangeReference)resolved!, context);
                        return true;

                    case NameShape.Refused:
                        operand = null!;
                        return false;

                    case NameShape.Scalar:
                        operand = new ScalarOperand(resolved!.Evaluate(context));
                        return true;

                    default:
                        operand = new ScalarOperand(expression.Evaluate(context));
                        return true;
                }
            }

            // ROW(x)/COLUMN(x) over a reference: ONE arm each, both walking the same shapes on their own
            // axis — see TryBuildPositionOperand, whose order Probe/ProbePosition mirrors exactly. The
            // pattern admits every reference NODE plus a defined name, and deliberately not a
            // reference-returning FUNCTION: see ResolvePositionRange.
            case Row { Arguments: [NameReference or Reference] } row:
                return TryBuildPositionOperand(
                    row,
                    row.Arguments[0],
                    PositionAxis.Row,
                    context,
                    out operand
                );

            case Column { Arguments: [NameReference or Reference] } column:
                return TryBuildPositionOperand(
                    column,
                    column.Arguments[0],
                    PositionAxis.Column,
                    context,
                    out operand
                );

            // The two lifted shapes, mirroring the Probe arms in the same order and on the same patterns.
            case UnaryOperation { Operator: not UnaryOperator.Plus } unary:
                return TryBuildUnary(unary, context, out operand);

            case Function function when TryGetLift(function, out var liftArguments):
                return TryBuildLift(function, liftArguments, context, out operand);

            case BinaryOperation binary:
                return TryBuildBinary(binary, context, out operand);

            case If ifNode when ifNode.Arguments.Length is 2 or 3:
                return TryBuildIf(ifNode, context, out operand);

            // The build twin of Probe's Let arm, in the same position — see TryBuildLet.
            case Let let:
                return TryBuildLet(let, context, out operand);

            // The build twin of Probe's producer arm, in the same position for the same reasons.
            case IArrayProducer producer:
                return producer.TryBuildArrayOperand(context, out operand);

            // Anything else is an opaque scalar: evaluate ONCE and broadcast. (This is where nested scalar
            // functions — SUM(A:A), a bare cell, a literal, a volatile RAND() — enter, without recursing.)
            default:
                operand = new ScalarOperand(expression.Evaluate(context));
                return true;
        }
    }

    // The build for ROW(x)/COLUMN(x) on one axis, walking the shapes in the SAME order ProbePosition does so
    // the two cannot drift: a rectangle written literally; one written INSIDE a shared-formula master (an
    // anchored node, resolved to its concrete delta-applied twin); the open-range refusal (the cost guard);
    // then anything that merely DENOTES a reference, through the oracle.
    //
    // The two rectangle arms are FAST PATHS, not requirements: they exist so the common case pays no
    // resolution at all. The ANCHORED one was a correctness fix when it was written (a slave's ROW(A1:A3) fell
    // to the opaque-scalar branch, and AnchoredRangeReference.Evaluate is always #VALUE! since a range has no
    // scalar value), but the oracle arm below now resolves an anchored range to the same delta-applied
    // rectangle. So no assertion in the suite distinguishes THE ANCHORED ARMS — this one and ProbePosition's
    // twin — from that fallback: deleting both leaves the whole suite green (re-verified, 1280/1280), because
    // what SharedFormulaSlaveFunctionTests pins is the delta, which the two routes apply alike.
    //
    // Neither rectangle arm is EQUIVALENT to the fallback, though, and one case tells them apart: the
    // fallback also runs ReferenceGuard.MissingSheet and degrades a reference on a DELETED sheet to Scalar,
    // where these arms hand back a plausible position vector for a sheet that no longer exists. That is the
    // same pre-existing gap ResolvePositionRange's comment records for a literal Ghost!A1:A3. On the LITERAL
    // arm an assertion DOES pin it —
    // MiniCseConsumerTests.Sum_OfRowOverLiteralRangeOnMissingSheet_KeepsTheSyntacticGap — which is no
    // contradiction of the paragraph above: the anchored twin of that case (a shared-formula slave whose
    // ROW argument names a deleted sheet) is simply not written, and it is the one assertion that would
    // make deleting the anchored arms visible.
    //
    // `node` is the ROW/COLUMN call and `argument` its single argument: Function declares no Arguments member,
    // so the caller — which has already destructured the argument to pattern-match it — passes both. The node
    // is needed only on the Scalar path, to evaluate itself once.
    //
    // The `default` arm below therefore only ever sees the caller's narrowed set (a defined name, a cell, a
    // union, a DynamicRange), NEVER a reference-returning function; widening either caller's pattern would
    // change that, which ResolvePositionRange explains is not free.
    private static bool TryBuildPositionOperand(
        Expression node,
        Expression argument,
        PositionAxis axis,
        EvaluationContext context,
        out ArrayOperand operand
    )
    {
        switch (argument)
        {
            case RangeReference range:
                operand = PositionOperand(range.GetBounds(), axis);
                return true;

            case AnchoredRangeReference anchored:
                operand = PositionOperand(anchored.ToRangeReference(context).GetBounds(), axis);
                return true;

            case OpenRangeReference:
                operand = null!;
                return false;

            default:
            {
                var shape = ResolvePositionRange(argument, context, out var bounds);

                if (shape is PositionArgumentShape.Refused)
                {
                    operand = null!;
                    return false;
                }

                operand =
                    shape is PositionArgumentShape.Array
                        ? PositionOperand(bounds, axis)
                        // A single cell, a union, or a name that resolves to no reference at all: let the
                        // node evaluate itself ONCE and broadcast, which is exactly its (correct) scalar
                        // answer, error included.
                        : new ScalarOperand(node.Evaluate(context));
                return true;
            }
        }
    }

    // The operand over one axis of a resolved rectangle: a VECTOR along that axis (an Nx1 column of row
    // numbers, a 1xM row of column numbers) whose other axis has extent 1 and therefore repeats under
    // broadcasting — Excel's shape, see PositionNumbersOperand.
    private static ArrayOperand PositionOperand(RangeBounds bounds, PositionAxis axis) =>
        axis is PositionAxis.Row
            ? new PositionNumbersOperand(bounds.TopRow, axis, bounds.RowCount, 1)
            : new PositionNumbersOperand(bounds.LeftColumn, axis, 1, bounds.ColumnCount);

    // What a ROW/COLUMN argument that DENOTES a reference contributes to the mini-CSE, once resolved.
    private enum PositionArgumentShape
    {
        // A rectangle: ROW/COLUMN over it is a vector of positions along its own axis — an Nx1 column of
        // row numbers, a 1xM row of column numbers (Aspose.Cells 26.6.0, measured 2026-09-10, CSE column:
        // COUNT(ROW(A1:C3)) is 3, not 9).
        Array,

        // No rectangle to spread over — a single cell, a union, a name that does not resolve: the ROW/COLUMN
        // node evaluates itself once and broadcasts its own scalar answer (or its own error).
        Scalar,

        // An open/whole-column reference: the cost guard refuses the whole array evaluation.
        Refused,
    }

    // Resolves a ROW/COLUMN argument to the rectangle whose positions the operand reports. This is the ONE
    // oracle Probe and TryBuildOperand share: the shape of such an argument is context-dependent, so a
    // syntactic probe could not answer it, and a probe that GUESSED (leaving the build to fail) would break
    // the documented "eligible iff the build succeeds — and the build is the SINGLE evaluation" contract.
    //
    // Because BOTH callers resolve, every argument admitted here is resolved TWICE per evaluation. For the
    // shapes the arms admit that is a scope/dictionary lookup and a virtual call — idempotent and cheap. It is
    // exactly why a reference-returning FUNCTION argument (ROW(INDEX(…))/OFFSET/INDIRECT) is NOT admitted:
    // resolving one evaluates the function's own arguments, so the pair of resolutions would draw a volatile
    // twice where the scalar path draws it once (measured: a counting function inside ROW(INDEX(…,TICK(),1))
    // is called exactly once, before and after this fix). Such an argument stays an opaque scalar here — still
    // its correct scalar answer (the resolved reference's top row / leftmost column), just not yet its array
    // shape.
    //
    // The one admitted shape whose resolution is not free is a DynamicRange with reference-returning ENDPOINTS
    // (ROW(INDEX(A1:A3,1,1):A3)), in scope deliberately as the same `or Reference` hook a structured table
    // column will land on. It costs no EXTRA draw: the scalar path this replaces already resolved it twice
    // (ReferenceGuard.MissingSheet, then ReferencePosition), so the count is 2 either way — measured. What a
    // SIDE-EFFECTING endpoint does expose is that the two resolutions can then disagree on the rectangle, and
    // the build's answer wins (ROW(INDEX(A1:A3,TICK(),1):A3) folds to 5, where Excel's single evaluation gives
    // 6 — and where the pre-fix scalar answer was a likewise-wrong 2). A pure endpoint, which is every real
    // one, resolves to the same rectangle twice and is exact.
    //
    // boundOpenRanges:false keeps an open range OPEN so it reaches the Refused arm instead of being silently
    // collapsed to its populated bounding box — both the cost guard this class states and the same choice
    // ROW/COLUMN's own scalar fallback (ReferencePosition) makes, where ROW(A:A) is the DECLARED row 1.
    private static PositionArgumentShape ResolvePositionRange(
        Expression argument,
        EvaluationContext context,
        out RangeBounds bounds
    )
    {
        bounds = default;

        if (
            !NamedReferences.TryResolveReference(
                argument,
                context,
                out var reference,
                boundOpenRanges: false
            )
        )
        {
            // Scalar, not Refused: the enclosing ROW/COLUMN then evaluates once and reports the ARGUMENT's own
            // error (#NAME? for an unknown name) through ReferencePosition. Refusing instead would hand the
            // consumer's scalar path an operand tree it never built, losing that error in an enclosing
            // operation (SUM(A1:A3+ROW(Nope)) must be #NAME?, not the #VALUE! of a range in a scalar add).
            return PositionArgumentShape.Scalar;
        }

        // A reference that resolves onto a DELETED sheet is a structural #REF!, which ROW/COLUMN's own
        // ReferenceGuard pass reports — degrade to Scalar so that error is what gets broadcast, rather than
        // handing back bounds and a plausible row number for a sheet that no longer exists. (The syntactic
        // arms above cannot reach this: a literal Ghost!A1:A3 carries its dead sheet name into the operand,
        // a pre-existing gap this fix neither widens nor closes.)
        if (ReferenceGuard.MissingSheet(reference, context) is not null)
        {
            return PositionArgumentShape.Scalar;
        }

        switch (reference)
        {
            case RangeReference range:
                // GetBounds() parses BOTH corners, so it is called once here and the rectangle handed back.
                bounds = range.GetBounds();
                return PositionArgumentShape.Array;

            case OpenRangeReference:
                return PositionArgumentShape.Refused;

            // A single cell (1x1 — nothing to spread over) or a union (no single position: #VALUE!).
            default:
                return PositionArgumentShape.Scalar;
        }
    }

    // What a bare defined name contributes to the mini-CSE, once resolved (Phase 11a Rule A).
    private enum NameShape
    {
        // Not a reference at all — a constant, a formula name that does NOT resolve to a reference, an
        // unknown name, a LET-bound scalar: the name evaluates itself once and broadcasts its own value (or
        // its own #NAME?). A formula name that DOES resolve to one is NOT opaque: NamedReferences
        // .TryResolveReference resolves the formula, so a name bound to `OFFSET(Sheet1!$A$1,0,0,3,1)` is
        // Range and streams the cells — SUM((OffName<>0)*1) = 2 here against 1 before Rule A, and 2 on the
        // oracle array-entered (#VALUE! typed; Aspose.Cells 26.6.0, measured 2026-09-10).
        Opaque,

        // A rectangle, on an existing sheet or not: the name streams the cells exactly like the literal.
        Range,

        // An open/whole-column reference: the cost guard refuses the whole array evaluation.
        Refused,

        // A single cell (1x1 — nothing to spread over) or a union (no single value: its own #VALUE!): the
        // RESOLVED node evaluates itself once and broadcasts.
        Scalar,
    }

    // Resolves a bare defined name to what it is bound to — the oracle Probe's and TryBuildOperand's
    // NameReference arms share, exactly as ResolvePositionRange is the one ROW/COLUMN share, and for the same
    // reason: the shape is context-dependent, and a probe that GUESSED would break the "eligible iff the
    // build succeeds" contract. Both callers resolve, so every name node is resolved twice per evaluation —
    // a scope/dictionary lookup and a virtual call, idempotent and cheap, the same double resolution
    // ResolvePositionRange documents. boundOpenRanges:false keeps an open range OPEN so it reaches Refused
    // instead of being silently collapsed to its populated bounding box (the cost guard).
    //
    // A rectangle whose SHEET IS MISSING is Range here, where ResolvePositionRange degrades it to Scalar.
    // The difference is what the operand carries (correction B2 of the sweep): a position vector had no
    // cell to read and would have handed back plausible row numbers for a sheet that no longer exists, so
    // ROW/COLUMN's own ReferenceGuard error is the right broadcast there; a VALUE vector must stream the
    // per-element #REF! that BuildRange yields, so that SUM((GhostName<>0)*1) is #REF! and
    // COUNT((GhostName<>"")*1) is 0 — what the literal Ghost!A1:A3 answers and what the oracle answers in
    // both entry modes (Aspose.Cells 26.6.0, 2026-09-10). Degrading to Scalar would broadcast ONE #REF!
    // value and make the COUNT a 1.
    private static NameShape ResolveNameShape(
        Expression expression,
        EvaluationContext context,
        out Reference? resolved
    )
    {
        if (
            !NamedReferences.TryResolveReference(
                expression,
                context,
                out resolved,
                boundOpenRanges: false
            )
        )
        {
            return NameShape.Opaque;
        }

        return resolved switch
        {
            RangeReference => NameShape.Range,
            OpenRangeReference => NameShape.Refused,
            _ => NameShape.Scalar,
        };
    }

    // Internal, not private, for the same reason Probe and TryBuildOperand are: a producer whose SOURCE
    // evaluated to a REFERENCE-kind scalar (INDIRECT("A1:A3"), OFFSET(A1,0,0,3,1), +A1:A3 — a reference-
    // returning node is an opaque scalar to the builder) resolves that rectangle through the one range
    // operand rather than wrapping the reference in a 1x1 singleton (SelectionProducers.TryBuildSource).
    internal static ArrayOperand BuildRange(RangeReference range, EvaluationContext context)
    {
        var workbook = context.Workbook;
        var handle = workbook.ResolveDenseHandle(range.SheetName);
        var bounds = range.GetBounds();

        return new RangeOperand(
            workbook,
            handle,
            range.SheetName,
            bounds.LeftColumn,
            bounds.TopRow,
            bounds.RowCount,
            bounds.ColumnCount
        );
    }

    private static bool TryBuildBinary(
        BinaryOperation binary,
        EvaluationContext context,
        out ArrayOperand operand
    )
    {
        if (
            !TryBuildOperand(binary.Left, context, out var left)
            || !TryBuildOperand(binary.Right, context, out var right)
        )
        {
            operand = null!;
            return false;
        }

        // Neither side is an array → the whole operation is a scalar (broadcast), evaluated once.
        if (!left.IsArray && !right.IsArray)
        {
            operand = new ScalarOperand(
                BinaryOperation.Apply(binary.Operator, left.Scalar, right.Scalar)
            );
            return true;
        }

        var (rows, columns) = ResultShape(left, right);
        operand = new BinaryOperand(binary.Operator, left, right, rows, columns);
        return true;
    }

    private static bool TryBuildIf(If ifNode, EvaluationContext context, out ArrayOperand operand)
    {
        if (!TryBuildOperand(ifNode.Arguments[0], context, out var condition))
        {
            operand = null!;
            return false;
        }

        // A scalar condition is not an element-wise selection: it picks ONE branch, which is then the IF's
        // whole operand. Only an ARRAY condition drives the zip below.
        if (!condition.IsArray)
        {
            return TryBuildScalarConditionIf(ifNode, condition.Scalar, context, out operand);
        }

        if (!TryBuildOperand(ifNode.Arguments[1], context, out var whenTrue))
        {
            operand = null!;
            return false;
        }

        // IF without an else branch yields a logical FALSE where the condition is false (Excel), which the
        // aggregators then ignore — the whole point of the SMALL(IF(...)) idiom.
        var hasElse = ifNode.Arguments.Length == 3;
        ArrayOperand whenFalse = new ScalarOperand(ComputedValue.Boolean(false));
        if (hasElse && !TryBuildOperand(ifNode.Arguments[2], context, out whenFalse))
        {
            operand = null!;
            return false;
        }

        // The IF's extent is the fold of all three operands (a scalar, the synthesized FALSE included,
        // contributes nothing): IF(E1:E3>1,E5:G5,0) is 3x3, not the condition's 3x1. Each operand then
        // projects that extent in its own At(), so a 2x1 condition over a 3x3 branch marks row 3 #N/A.
        var shape = new ShapeFold();
        shape.Fold(condition);
        shape.Fold(whenTrue);
        shape.Fold(whenFalse);

        operand = new IfOperand(condition, whenTrue, whenFalse, shape.Rows, shape.Columns);
        return true;
    }

    // IF under a SCALAR condition: the condition is coerced exactly as If.Evaluate coerces it (the text
    // "TRUE"/"FALSE" included), only the taken branch is touched (short-circuit), and that branch's operand
    // IS the IF's operand — so a producer or any other computed array in the branch streams whole, where
    // the former "opaque scalar" answer (ScalarOperand(ifNode.Evaluate)) collapsed it to its top-left
    // element: SUM(IF(TRUE,SEQUENCE(3),0)) answered 1 for the oracle's 6, silently, because a producer's
    // scalar rule is FirstElement. `conditionValue` is the value the caller already built, so the condition
    // is evaluated once (the former path evaluated it a second time inside If.Evaluate).
    //
    // THE BUILD NEVER RETURNS false ON THE ELIGIBLE PATH, AND THAT IS THE WHOLE POINT — an earlier version
    // DECLINED for a taken bare-reference branch, handing the choice back to the consumer's scalar path,
    // which re-entered If.Evaluate and evaluated the condition a SECOND time. With a volatile condition the
    // two draws disagree and the fallback then takes the OTHER branch as a scalar, so a producer collapsed to
    // its top-left again: over 400 seeded workbooks SUM(IF(RAND()<0.5,A1:A3,SEQUENCE(3))) answered 1 in
    // exactly 100 of them — the very silent wrong number this method exists to remove. Found by the phase's
    // final review, which tested the MIXED shape where an earlier check had tested only two producers, the
    // one shape that never reaches the decline.
    //
    // So every path now returns true with an operand, and the Probe's promise (IsArrayEligible == the build
    // yields an array) holds in three cases rather than two:
    //   - NOT array-eligible (no computed-array branch — ProbeIfBranches): the taken branch is evaluated as
    //     the scalar If.Evaluate would have produced, unchanged from before.
    //   - Eligible, and the taken branch builds as an array: that operand IS the IF's operand.
    //   - Eligible, and the taken branch yields a SCALAR — IF(FALSE,SEQUENCE(3),0), an error condition, a
    //     branch-less IF's FALSE, or a bare reference — the scalar is handed back through WrapScalar as a 1x1
    //     array, never a ScalarOperand, so the consumer that probed "array" cannot fall back and re-draw.
    //     A 1x1 broadcasts like the scalar it holds, so no consumer reads it differently.
    //
    // WrapScalar is what keeps a bare-reference branch answering what it answered before, WITHOUT deciding
    // the "IF returns a reference" question: it reproduces the scalar path's own answer once instead of
    // twice, and the asymmetry between a name and a bare range is not a choice made here — it falls out of
    // what each node's Evaluate returns. RangeReference.Evaluate is #VALUE! (a plain error, so a 1x1 error),
    // while NameReference.Evaluate carries a reference value, which WrapScalar resolves through BuildRange
    // exactly as SelectionProducers.TryBuildSource already does for INDIRECT/OFFSET. Measured on this build:
    // SUM(IF(TRUE,A1:A3,0)) and SUM(IF(TRUE,A1:A3,SEQUENCE(3))) are both #VALUE! and ROWS of the second is
    // #VALUE! (the oracle says 14, 14 and 3 — sweep item 32, unmoved); SUM(IF(TRUE,MyName,0)) and
    // SUM(IF(TRUE,MyName,SEQUENCE(3))) are both 14; an OPEN range resolves to the loud 1x1 #VALUE! its
    // shape always gets (SUM(IF(TRUE,MyColumn,SEQUENCE(3)))), and in the UNTAKEN branch it still costs
    // nothing (SUM(IF(FALSE,MyColumn,0)*B1:B3) = 0, SUM(IF(TRUE,SEQUENCE(3),MyColumn)) = 6).
    // One row MOVED with the fix and moved toward the oracle: a single-cell name in the mixed shape,
    // SUM(IF(TRUE,MyCell,SEQUENCE(3))) over a text cell, went #VALUE! -> 0, which is the oracle's answer.
    // Its plain twin SUM(IF(TRUE,MyCell,0)) stays #VALUE! because it is not array-eligible at all.
    //
    // WHAT THIS METHOD DOES NOT PROMISE. A gate that keys on the PROBE rather than on the built operand sees
    // "array" and refuses before any of the above runs, so a bare-reference branch is NOT interchangeable
    // with a scalar sibling for those consumers. Measured, and pinned: COUNTIF(IF(TRUE,A1:A3,SEQUENCE(3)),
    // ">0") is #REF! while COUNTIF(IF(TRUE,A1:A3,0),">0") is 0 (the oracle answers 2 for both), and
    // COUNTBLANK of the same pair is #REF! against 0 (the oracle answers 1). An earlier version of this
    // comment claimed the decline "answers exactly what it answers with a scalar sibling"; for the criteria
    // family and COUNTBLANK that was false, and it is the probe, not the build, that makes it so.
    private static bool TryBuildScalarConditionIf(
        If ifNode,
        ComputedValue conditionValue,
        EvaluationContext context,
        out ArrayOperand operand
    )
    {
        var isArray = ProbeIfBranches(ifNode, context, conditionIsArray: false).IsArray;

        if (conditionValue.CoerceToBoolAllowingTextWords(out var taken) is { } error)
        {
            operand = Wrap(ComputedValue.Error(error), isArray);
            return true;
        }

        if (!taken && ifNode.Arguments.Length == 2)
        {
            operand = Wrap(ComputedValue.Boolean(false), isArray);
            return true;
        }

        var branch = ifNode.Arguments[taken ? 1 : 2];

        if (!isArray)
        {
            operand = new ScalarOperand(branch.Evaluate(context));
            return true;
        }

        if (IsBareReferenceNode(branch, context))
        {
            operand = WrapScalar(branch.Evaluate(context), context);
            return true;
        }

        if (!TryBuildOperand(branch, context, out var built))
        {
            operand = null!;
            return false;
        }

        operand = built.IsArray ? built : WrapScalar(built.Scalar, context);
        return true;

        static ArrayOperand Wrap(ComputedValue value, bool asArray) =>
            asArray ? new SingletonArrayOperand(value) : new ScalarOperand(value);

        static ArrayOperand WrapScalar(ComputedValue value, EvaluationContext context) =>
            value.TryGetReference(out var reference)
                ? reference switch
                {
                    RangeReference range => BuildRange(range, context),
                    CellReference cell => new SingletonArrayOperand(cell.Evaluate(context)),
                    _ => new SingletonArrayOperand(ComputedValue.Error(Error.Value)),
                }
                : new SingletonArrayOperand(value);
    }

    // The shape of a binary result: a scalar takes the other side's shape; two arrays fold per axis by
    // Broadcasting.Axis (an extent of 1 defers to the other side, two larger extents take the maximum), and
    // each operand then answers #N/A at the positions of that extent it does not cover. Expressed as two
    // steps of the N-ary ShapeFold so the lift and the binary operation share ONE rule.
    private static (int Rows, int Columns) ResultShape(ArrayOperand left, ArrayOperand right)
    {
        var shape = new ShapeFold();
        shape.Fold(left);
        shape.Fold(right);

        return (shape.Rows, shape.Columns);
    }

    // The running result shape of an N-ary element-wise node: a scalar contributes nothing; the first array
    // sets the shape; every further array folds in per axis through Broadcasting.Axis — an extent of 1
    // defers to the other operand, two larger extents take the maximum. "No array seen yet" is an explicit
    // flag rather than a (0, 0) sentinel because Axis(0, 1) is 0: a (0, 0) accumulator folding a 1xN first
    // operand would yield a 0-row extent — an empty stream, SUM = 0, silently. (Phase 10 also assumed the
    // flag would one day carry a legitimate 0-row array from an empty FILTER; Phase 7 forbids that shape
    // instead — ArrayShaping's invariant makes an empty producer result a 1x1 singleton, so this fold does
    // not expect a 0 extent from any operand; ArrayProducerContractTests shows what a 0x1 would do.)
    //
    // Which positions of the folded extent an operand covers is decided at READ time by that operand's own
    // At() through Broadcasting.TryProject: an uncovered position answers #N/A there, so the node's body
    // receives a real per-element error, which an error-propagating body (arithmetic, LEN, ROUND, …) carries
    // to that element and an error-consuming body (IFERROR, IS*, N, T, IFS, SWITCH) recovers —
    // SUM(IFERROR(A1:C3*H1:H2,0)) is 36 (VectorBroadcastingTests, Aspose.Cells 26.6.0, 2026-09-10, CSE).
    private struct ShapeFold
    {
        private bool _seen;

        public int Rows { get; private set; }
        public int Columns { get; private set; }

        public void Fold(ArrayOperand operand)
        {
            if (!operand.IsArray)
            {
                return;
            }

            if (!_seen)
            {
                _seen = true;
                Rows = operand.Rows;
                Columns = operand.Columns;
                return;
            }

            Rows = Broadcasting.Axis(Rows, operand.Rows);
            Columns = Broadcasting.Axis(Columns, operand.Columns);
        }
    }

    // Whether `function` is a registered built-in the mini-CSE may lift, handing back its arguments. The
    // registry lookup is the only way to reach a Function's arguments (Function declares no Arguments
    // member) and to rebuild the node (entry.Create), so the classification rides along for free; it is paid
    // once per function NODE per evaluation, never per element. `Length > 0` keeps a zero-argument entry
    // (PI, RAND, …) out of the arm without a special case; a custom FunctionCall is not in the table.
    private static bool TryGetLift(Function function, out Expression[] arguments)
    {
        if (
            FunctionRegistry.ByType.TryGetValue(function.GetType(), out var entry)
            && entry.Lifting is ArrayLifting.Elementwise
        )
        {
            arguments = entry.GetArguments(function);
            return arguments.Length > 0;
        }

        arguments = [];
        return false;
    }

    // THE decision both lifted arms share, in Probe and in the build alike: a lifted shape is an array when
    // every argument's build would succeed and at least one is an array. A refused argument (an open range
    // somewhere below) does NOT refuse the shape — it makes it an OPAQUE SCALAR, evaluated once, exactly as
    // the pre-lift `default` arm treated the whole node. Refusing instead would unwind an enclosing array
    // expression that tolerates the open range today (measured: SUM(IF(A1:A3>0,1,LEN(B:B))) = 3 must stay 3,
    // where a refusal turns it into the scalar path's #VALUE!). It never evaluates: only the classification
    // and the arguments' own probes are consulted.
    //
    // The builders call it BEFORE building any operand, which is what keeps every scalar argument evaluated
    // exactly once: building operands first and only then discovering that the shape is a scalar (or that a
    // later argument refuses) would evaluate the built scalars once in their ScalarOperand and once more in
    // the node's own Evaluate — a volatile drawn twice where the pre-lift path drew it once. (A span, so the
    // unary arms' single-operand collection expression is stack-allocated.)
    private static bool ProbeLift(ReadOnlySpan<Expression> arguments, EvaluationContext context)
    {
        var isArray = false;

        foreach (var argument in arguments)
        {
            var (succeeds, argumentIsArray) = Probe(argument, context);

            if (!succeeds)
            {
                return false;
            }

            isArray |= argumentIsArray;
        }

        return isArray;
    }

    private static bool TryBuildUnary(
        UnaryOperation unary,
        EvaluationContext context,
        out ArrayOperand operand
    )
    {
        // Not an array (a scalar operand, or a refused one): the whole unary is an opaque scalar, evaluated
        // ONCE through the node's own path — Probe answered (true, false) for it, so the build must succeed
        // as a non-array.
        if (!ProbeLift([unary.Operand], context))
        {
            operand = new ScalarOperand(unary.Evaluate(context));
            return true;
        }

        if (!TryBuildOperand(unary.Operand, context, out var inner))
        {
            operand = null!;
            return false;
        }

        operand = new UnaryOperand(unary.Operator, inner, inner.Rows, inner.Columns);
        return true;
    }

    private static bool TryBuildLift(
        Function function,
        Expression[] arguments,
        EvaluationContext context,
        out ArrayOperand operand
    )
    {
        // Same opaque-scalar fallback as TryBuildUnary: today's behaviour, unchanged, evaluated once.
        if (!ProbeLift(arguments, context))
        {
            operand = new ScalarOperand(function.Evaluate(context));
            return true;
        }

        var operands = new ArrayOperand[arguments.Length];
        var shape = new ShapeFold();

        for (var i = 0; i < arguments.Length; i++)
        {
            if (!TryBuildOperand(arguments[i], context, out operands[i]))
            {
                operand = null!;
                return false;
            }

            shape.Fold(operands[i]);
        }

        operand = new LiftedFunctionOperand(
            FunctionRegistry.ByType[function.GetType()].Create,
            arguments,
            operands,
            context,
            shape.Rows,
            shape.Columns
        );
        return true;
    }

    /// <summary>
    /// A no-alloc, lazy view over the mini-CSE array: the resolved operand tree plus its row-major shape.
    /// <see cref="ElementAt"/> computes a single element on demand (INDEX walks to the n-th only); a
    /// <c>foreach</c> binds the struct <see cref="Enumerator"/> by duck typing (the <see cref="List{T}"/>
    /// pattern), so the aggregating consumers (SUM/COUNT/AVERAGE/MIN/MAX, SMALL/LARGE) scan every element
    /// row-major with no interface dispatch and no <c>ComputedValue[]</c>.
    /// </summary>
    internal readonly struct ArrayStream
    {
        private readonly ArrayOperand _root;

        internal ArrayStream(ArrayOperand root, int rows, int columns)
        {
            _root = root;
            Rows = rows;
            Columns = columns;
        }

        public int Rows { get; }
        public int Columns { get; }
        public int Length => Rows * Columns;

        /// <summary>The value at a row-major index (0-based), computed on demand.</summary>
        public ComputedValue ElementAt(int index) => _root.At(index, Rows, Columns);

        public Enumerator GetEnumerator() => new(_root, Rows, Columns);

        public struct Enumerator
        {
            private readonly ArrayOperand _root;
            private readonly int _rows;
            private readonly int _columns;
            private readonly int _length;
            private int _index;
            private ComputedValue _current;

            internal Enumerator(ArrayOperand root, int rows, int columns)
            {
                _root = root;
                _rows = rows;
                _columns = columns;
                _length = rows * columns;
                _index = -1;
                _current = default;
            }

            public readonly ComputedValue Current => _current;

            public bool MoveNext()
            {
                var next = _index + 1;
                if (next >= _length)
                {
                    return false;
                }

                _index = next;
                _current = _root.At(next, _rows, _columns);
                return true;
            }
        }
    }
}
