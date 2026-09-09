using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

/// <summary>
/// AGGREGATE — Microsoft documents it as TWO syntaxes over one function name:
/// <c>AGGREGATE(function_num, options, ref1, [ref2], …)</c> (the reference form) and
/// <c>AGGREGATE(function_num, options, array, [k])</c> (the array form).
///
/// <para>Nothing SYNTACTIC separates them — <c>AGGREGATE(9,6,A1:A3,B1:B3)</c> and
/// <c>AGGREGATE(15,6,A1:A3,2)</c> are both four-argument calls — so the form is chosen by
/// <c>function_num</c> alone, at evaluation time: 1-13 are the whole-population folds (the reference form,
/// every argument from index 2 on is a ref), 14-19 are the positional selections (the array form, whose
/// fourth argument is k / quart). Hence the registry entry's <c>MaxArgs = int.MaxValue</c> and the array
/// form's own "exactly 4 arguments" rule below, which is where the documented "if a second ref argument is
/// necessary but not provided, AGGREGATE returns a #VALUE! error" lives.</para>
///
/// <para>The scan, the accumulator and both halves of the function_num map live in
/// <see cref="AggregateCodes"/>, shared with <see cref="Subtotal"/>; this node is Subtotal's shape plus the
/// form selection and the options bits.</para>
/// </summary>
[MemoryPackable]
public sealed partial record Aggregate(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var rawCode) is { } codeError)
        {
            return ComputedValue.Error(codeError);
        }

        var code = (int)Math.Truncate(rawCode);

        if (code is < 1 or > 19)
        {
            return ComputedValue.Error(Error.Value);
        }

        // An OMITTED options argument arrives as BlankValue (Parser.ParseArgument) and coerces to 0 —
        // Excel's documented "0 or omitted".
        if (Arguments[1].Evaluate(context).CoerceToNumber(out var rawOptions) is { } optionsError)
        {
            return ComputedValue.Error(optionsError);
        }

        var options = (int)Math.Truncate(rawOptions);

        if (options is < 0 or > 7)
        {
            return ComputedValue.Error(Error.Value);
        }

        // The options table decomposes into three independent bits:
        //   bit0 (1) = ignore hidden rows   — a NO-OP here: MySheet has no hidden-row model, so options
        //                                     1/3/5/7 behave exactly as 0/2/4/6. This is the documented
        //                                     model limit (the plan's S6 caveat); the bit is read by
        //                                     NOBODY on purpose, rather than silently rejected.
        //   bit1 (2) = ignore error values  — options 2, 3, 6, 7.
        //   bit2 (4) = do NOT ignore nested SUBTOTAL/AGGREGATE — the nested skip is documented for options
        //                                     0-3 ("Ignore nested SUBTOTAL and AGGREGATE functions"), and
        //                                     4-7 ("Ignore nothing" / hidden rows / errors) keep them.
        var ignoreErrors = (options & 2) != 0;
        var skip =
            (options & 4) == 0
                ? AggregateCodes.NestedSkip.SubtotalAndAggregate
                : AggregateCodes.NestedSkip.None;

        return code <= 13
            ? ReferenceForm(code, ignoreErrors, skip, context)
            : ArrayForm(code, ignoreErrors, skip, context);
    }

    // AGGREGATE(function_num, options, ref1, [ref2], …) for function_num 1-13 — Subtotal.Evaluate's tail
    // verbatim, with the two option flags threaded through instead of hard-coded.
    private ComputedValue ReferenceForm(
        int code,
        bool ignoreErrors,
        AggregateCodes.NestedSkip skip,
        EvaluationContext context
    )
    {
        var refs = Arguments[2..];

        // A reference argument to a missing sheet is a structural #REF! that short-circuits the whole
        // function, before any per-cell scan. Load-bearing under the error-ignoring options: without it a
        // missing-sheet range would aggregate to an EMPTY population and answer 0 (or #NUM!) instead.
        if (ReferenceGuard.MissingSheet(refs, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var accumulator = new AggregateCodes.Accumulator(code, ignoreErrors);

        foreach (var argument in refs)
        {
            if (AggregateCodes.Feed(argument, context, ref accumulator, skip) is { } error)
            {
                return ComputedValue.Error(error);
            }
        }

        return accumulator.Finish();
    }

    // AGGREGATE(function_num, options, array, k) for function_num 14-19.
    private ComputedValue ArrayForm(
        int code,
        bool ignoreErrors,
        AggregateCodes.NestedSkip skip,
        EvaluationContext context
    )
    {
        // "If a second ref argument is necessary but not provided, AGGREGATE returns a #VALUE! error" —
        // 14-19 all take a k / quart, so the array form is exactly four arguments.
        if (Arguments.Length != 4)
        {
            return ComputedValue.Error(Error.Value);
        }

        var array = Arguments[2];

        if (ReferenceGuard.MissingSheet(array, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        // The mini-CSE gate, identical to AggregateCodes.Feed's (and to OrderSelection.KthValue's): a
        // non-Reference argument the mini-CSE can evaluate element-wise is STREAMED, while a plain range
        // keeps the cell-by-cell scan — which is what carries the nested-aggregate skip. Feed itself cannot
        // be reused here because 14/15 need the raw stream, not an accumulator: the bounded heap selects
        // the k-th value without ever materializing (or sorting) the vector.
        if (
            array is not Reference
            && ArrayEvaluation.IsArrayEligible(array, context)
            && ArrayEvaluation.TryEvaluateStream(array, context, out var stream)
        )
        {
            if (code is 14 or 15)
            {
                return OrderSelection.KthValueStreaming(
                    stream,
                    Arguments[3],
                    context,
                    largest: code == 14,
                    ignoreErrors
                );
            }

            // 16-19 (the percentiles and quartiles) need the WHOLE population anyway, so the heap buys
            // nothing: collect the stream through the same accumulator the reference path uses.
            var streamed = new AggregateCodes.Accumulator(code, ignoreErrors);

            if (AggregateCodes.CollectStream(stream, ref streamed) is { } streamError)
            {
                return ComputedValue.Error(streamError);
            }

            return Select(code, streamed, context);
        }

        var accumulator = new AggregateCodes.Accumulator(code, ignoreErrors);

        if (AggregateCodes.Gather(array, context, ref accumulator, skip) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return Select(code, accumulator, context);
    }

    // AggregateCodes.Positional requires an ASCENDING population, and the accumulator gathers in scan
    // order, so the sort happens HERE — Fold's in-place sort for code 12 is not on this path at all.
    private ComputedValue Select(
        int code,
        AggregateCodes.Accumulator accumulator,
        EvaluationContext context
    )
    {
        var numbers = accumulator.Numbers;
        numbers.Sort();

        return AggregateCodes.Positional(code, numbers, Arguments[3], context);
    }
}
