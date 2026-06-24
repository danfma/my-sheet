using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record SumOperator(Expression[] Expressions) : Operator
{
    public override object? Compute(Workbook workbook)
    {
        var result = 0.0;

        foreach (var expression in Expressions)
        {
            switch (expression)
            {
                case NumberValue number:
                    result += number.Value;
                    break;

                case StringValue stringValue
                    when double.TryParse(stringValue.Value, out var parsedDouble):
                    result += parsedDouble;
                    break;

                case StringValue:
                    return ErrorValue.NotValue;

                case BooleanValue booleanValue:
                    result += booleanValue.Value ? 1 : 0;
                    break;

                case CellReference cellReference:
                    var resolved = cellReference.Resolve(workbook);

                    if (IsNumberReference(resolved, out var resolvedNumber))
                    {
                        result += resolvedNumber;
                    }
                    else if (resolved is StringValue or BooleanValue or BlankValue)
                    {
                        continue;
                    }
                    else if (resolved is ErrorValue)
                    {
                        return resolved;
                    }
                    else
                    {
                        return ErrorValue.NotValue;
                    }
                    break;

                default:
                    throw new NotImplementedException();
            }
        }

        return result;

        static bool IsNumberReference(Expression? expression, out double number)
        {
            if (expression is NumberValue numberValue)
            {
                number = numberValue.Value;

                return true;
            }

            number = 0;

            return false;
        }
    }
}