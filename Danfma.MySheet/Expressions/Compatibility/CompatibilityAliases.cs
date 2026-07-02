using Danfma.MySheet.Expressions.Statistical;
using MemoryPack;

namespace Danfma.MySheet.Expressions.Compatibility;

// The pre-2010 legacy names of the modern statistical functions. Each is a DISTINCT record (never
// the modern node): the un-parse (FormulaWriter.Call) renders the node's own name, so a formula
// entered as STDEV(...) survives FORMULATEXT/serialization/export as STDEV(...) instead of being
// rewritten to STDEV.S(...). The behaviour is shared with the modern node via its internal static
// Compute method.

/// <summary>Legacy alias of MODE.SNGL.</summary>
[MemoryPackable]
public sealed partial record Mode(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ModeSngl.Compute(Arguments, context);
}

/// <summary>Legacy alias of STDEV.S.</summary>
[MemoryPackable]
public sealed partial record StDev(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        StDevS.Compute(Arguments, context);
}

/// <summary>Legacy alias of STDEV.P.</summary>
[MemoryPackable]
public sealed partial record StDevP(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Statistical.StDevP.Compute(Arguments, context);
}

/// <summary>Legacy alias of VAR.S.</summary>
[MemoryPackable]
public sealed partial record Var(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        VarS.Compute(Arguments, context);
}

/// <summary>Legacy alias of VAR.P.</summary>
[MemoryPackable]
public sealed partial record VarP(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Statistical.VarP.Compute(Arguments, context);
}

/// <summary>Legacy alias of RANK.EQ.</summary>
[MemoryPackable]
public sealed partial record Rank(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        RankEq.Compute(Arguments, context, average: false);
}

/// <summary>Legacy alias of PERCENTILE.INC.</summary>
[MemoryPackable]
public sealed partial record Percentile(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        PercentileInc.Compute(Arguments, context);
}

/// <summary>Legacy alias of PERCENTRANK.INC.</summary>
[MemoryPackable]
public sealed partial record PercentRank(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        PercentRankInc.Compute(Arguments, context, exclusive: false);
}

/// <summary>Legacy alias of QUARTILE.INC.</summary>
[MemoryPackable]
public sealed partial record Quartile(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        QuartileInc.Compute(Arguments, context);
}

/// <summary>Legacy alias of COVARIANCE.P.</summary>
[MemoryPackable]
public sealed partial record Covar(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        CovarianceP.Compute(Arguments, context);
}

/// <summary>Legacy alias of FORECAST.LINEAR.</summary>
[MemoryPackable]
public sealed partial record Forecast(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ForecastLinear.Compute(Arguments, context);
}
