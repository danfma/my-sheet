using MemoryPack;

namespace Danfma.MySheet.Expressions;

/// <summary>A rectangular, row-major Excel array constant containing literal values only.</summary>
[MemoryPackable]
public sealed partial record ArrayConstant(Expression[] Values, int Rows, int Columns)
    : Expression,
        IArrayProducer
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ArrayEvaluation.FirstElement(this, context);

    (bool Succeeds, bool IsArray) IArrayProducer.ProbeArray(EvaluationContext context) =>
        (true, true);

    bool IArrayProducer.TryBuildArrayOperand(EvaluationContext context, out ArrayOperand operand)
    {
        var values = new ComputedValue[Values.Length];
        for (var i = 0; i < Values.Length; i++)
        {
            values[i] = Values[i].Evaluate(context);
        }

        operand = new ArrayConstantOperand(values, Rows, Columns);
        return true;
    }
}

internal sealed class ArrayConstantOperand : ArrayOperand
{
    private readonly ComputedValue[] _values;
    private readonly int _rows;
    private readonly int _columns;

    public ArrayConstantOperand(ComputedValue[] values, int rows, int columns)
    {
        ArrayShaping.RequireProducerShape(rows, columns);
        _values = values;
        _rows = rows;
        _columns = columns;
    }

    public override bool IsArray => true;
    public override int Rows => _rows;
    public override int Columns => _columns;

    public override ComputedValue At(int index, int rows, int columns) =>
        Broadcasting.TryProject(index, rows, columns, _rows, _columns, out var own)
            ? _values[own]
            : ComputedValue.Error(Error.NA);
}
