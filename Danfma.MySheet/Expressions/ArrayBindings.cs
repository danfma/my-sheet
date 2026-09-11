namespace Danfma.MySheet.Expressions;

/// <summary>
/// What a BINDING SITE captures (Phase 11c, array bindings): a <c>LET</c> binding today, and — Task 3 of the
/// same phase — <c>CHOOSE</c>'s chosen branch, the operand of unary <c>+</c> and a defined name's definition.
/// One rule for all of them, <see cref="Capture"/>: an expression the mini-CSE would stream as an array — not
/// a bare reference node, array-eligible in the binding's own scope — is built ONCE, here, and the OPERAND
/// is what gets bound; anything else goes through <see cref="NamedReferences.CaptureValue"/> exactly as
/// before (a range node stays a reference value, a scalar evaluates). The operand lives for the scope's
/// lifetime — one evaluation — which is what makes evaluate-once fall out by construction:
/// <c>LET(x,SEQUENCE(3,1,RAND(),0),SUM(x)-SUM(x))</c> is 0 because the second read is the same operand, not
/// a second build (<c>ArrayBindingTests</c> pins it with a counting custom function).
///
/// <para>The SCALAR reading of an array binding is the operand's top-left (<see cref="Binding.TopLeft"/>,
/// through <see cref="ArrayEvaluation.FirstElement(ArrayOperand)"/>) — the same <c>@</c> rule a bare producer
/// follows, and the oracle's answer for an operator binding read bare: <c>=LET(x,A1:A3*2,x)</c> is 10 and
/// <c>=LET(x,IF(A1:A3&gt;0,A1:A3),x)</c> is 5 (Aspose.Cells 26.6.0, 2026-09-11, array-entered column; plain
/// entry intersects per row), where the scalar path's own answer for either expression is <c>#VALUE!</c>.
/// <see cref="NameReference.Evaluate"/> is where that reading happens for a LET name.</para>
/// </summary>
internal static class ArrayBindings
{
    /// <summary>
    /// A captured binding: an array <see cref="Operand"/>, or the <see cref="Value"/>
    /// <see cref="NamedReferences.CaptureValue"/> gave — never both. <see cref="EvaluationContext.WithName(string, Binding)"/>
    /// binds whichever it holds.
    /// </summary>
    internal readonly struct Binding
    {
        public ArrayOperand? Operand { get; }
        public ComputedValue Value { get; }

        public Binding(ArrayOperand operand)
        {
            Operand = operand;
            Value = default;
        }

        public Binding(ComputedValue value)
        {
            Operand = null;
            Value = value;
        }

        public bool IsArray => Operand is not null;

        /// <summary>The scalar reading: the operand's top-left element, or the captured value itself.</summary>
        public ComputedValue TopLeft =>
            Operand is { } operand ? ArrayEvaluation.FirstElement(operand) : Value;
    }

    /// <summary>
    /// Captures <paramref name="expression"/> for binding — its SINGLE evaluation. The gate is
    /// <see cref="ArrayEvaluation.TryStream"/>'s own, in its order: a bare reference node (the context-aware
    /// <see cref="ArrayEvaluation.IsBareReferenceNode(Expression, EvaluationContext)"/>, so a name already
    /// bound to an array is NOT bare and rebinds its operand — <c>LET(f,FILTER(…),g,f,SUM(g))</c>) keeps
    /// <see cref="NamedReferences.CaptureValue"/>'s reference path, a range binding staying a reference value
    /// (Phase 11a Rule A); the probe never evaluates; and once the probe has said yes the build is that one
    /// evaluation and yields an array by the probe's contract (<see cref="ArrayEvaluation.IsArrayEligible"/>
    /// iff the build yields an array). The fall-through after an eligible probe is therefore unreachable —
    /// were it ever reached, <see cref="NamedReferences.CaptureValue"/> would evaluate the expression a
    /// second time, the double draw the contract exists to prevent, so a failure there is a bug in
    /// <see cref="ArrayEvaluation.Probe"/>/<see cref="ArrayEvaluation.TryBuildOperand"/>, not here.
    /// </summary>
    public static Binding Capture(Expression expression, EvaluationContext context)
    {
        if (
            !ArrayEvaluation.IsBareReferenceNode(expression, context)
            && ArrayEvaluation.IsArrayEligible(expression, context)
            && ArrayEvaluation.TryBuildOperand(expression, context, out var operand)
            && operand.IsArray
        )
        {
            return new Binding(operand);
        }

        return new Binding(NamedReferences.CaptureValue(expression, context));
    }

    /// <summary>
    /// The probe's stand-in for <see cref="Capture"/>: the SHAPE the binding would have, discovered WITHOUT
    /// evaluating anything, so <see cref="ArrayEvaluation.Probe"/> can walk a <c>LET</c>'s bindings and stay
    /// the type-walk <see cref="ArrayEvaluation.IsArrayEligible"/> promises. Same gate, same order: an
    /// array-eligible expression binds <see cref="ShapeOnlyOperand"/> — an operand that answers
    /// <see cref="ArrayOperand.IsArray"/> and nothing else — and anything else binds the reference it RESOLVES
    /// to (<see cref="NamedReferences.TryResolveReference"/> with open ranges kept open, so the cost guard still
    /// sees them: a range node, a name bound to one, a reference-returning function such as <c>OFFSET</c>) or
    /// an opaque blank, which the <see cref="NameReference"/> arms read as exactly the shape the captured
    /// value would have — a scalar's own value never matters to a probe. This is the same resolution
    /// <c>ResolveNameShape</c> applies to a formula-defined name, with the same blind spot: a node that passes
    /// a child's reference value out of <c>Evaluate</c> without answering <c>TryResolveReference</c> itself
    /// (<c>IF(TRUE,MyName,0)</c> over a range-bound name) is an opaque scalar to the probe and a range-bound
    /// name to the build. The probe is then MORE conservative than the build, never the reverse, so a
    /// consumer that probed "not an array" keeps the scalar path it takes today — measured and pinned in
    /// <c>ArrayBindingTests</c> (<c>AChainedRebindingOfAnIfOverARangeName_…</c>): <c>SUM(LET(a,IF(TRUE,
    /// Rng,0),b,a,b*1))</c> stays <c>#VALUE!</c> and <c>COUNTIF</c> of the same Let stays 0 (the oracle
    /// answers 14, and <c>#REF!</c> array-entered), while the build side answers the oracle's 14 for
    /// <c>LET(a,IF(TRUE,Rng,0),b,a,SUM(b*1))</c> and 32 for the Let under <c>*B1:B3</c>. The seam closes
    /// with the "IF returns a reference" decision (sweep item 32).
    /// </summary>
    public static Binding Shape(Expression expression, EvaluationContext context)
    {
        if (
            !ArrayEvaluation.IsBareReferenceNode(expression, context)
            && ArrayEvaluation.IsArrayEligible(expression, context)
        )
        {
            return new Binding(ShapeOnlyOperand.Instance);
        }

        return new Binding(
            NamedReferences.TryResolveReference(
                expression,
                context,
                out var reference,
                boundOpenRanges: false
            )
                ? ComputedValue.Reference(reference)
                : ComputedValue.Blank
        );
    }

    // The probe's array stand-in: IsArray is the only thing a probe reads off a binding. Its extent and
    // elements do not exist — a read means a probe path evaluated a bound name, which is a bug — so At
    // throws rather than answering something plausible.
    private sealed class ShapeOnlyOperand : ArrayOperand
    {
        public static readonly ShapeOnlyOperand Instance = new();

        private ShapeOnlyOperand() { }

        public override bool IsArray => true;
        public override int Rows => 0;
        public override int Columns => 0;

        public override ComputedValue At(int index, int rows, int columns) =>
            throw new InvalidOperationException(
                "The probe's shape-only array binding has no elements: a probe path evaluated a LET-bound name."
            );
    }
}
