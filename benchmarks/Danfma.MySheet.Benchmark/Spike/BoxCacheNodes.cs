namespace Danfma.MySheet.Benchmark.Spike;

// Variante C — plano-B: mantém o retorno `object?`, mas evita alocar boxes para os valores comuns
// (inteiros pequenos não-negativos e true/false). É a mitigação barata que NÃO exige migrar os 64
// arquivos. O benchmark mostra quanto da alocação ela remove: ajuda os literais pequenos, mas os
// resultados que compõem (somas grandes) continuam alocando — ao contrário do CellValue, que zera tudo.

/// <summary>Cache de boxes para inteiros 0..255 e para os dois booleanos (padrão do runtime para ints).</summary>
internal static class BoxCache
{
    private static readonly object[] SmallInts = BuildSmallInts();
    private static readonly object BoxedTrue = true;
    private static readonly object BoxedFalse = false;

    private static object[] BuildSmallInts()
    {
        var boxes = new object[256];
        for (var i = 0; i < boxes.Length; i++)
        {
            boxes[i] = (double)i;
        }

        return boxes;
    }

    public static object Number(double n)
    {
        if (n is >= 0d and < 256d)
        {
            var i = (int)n;
            if (i == n)
            {
                return SmallInts[i]; // sem alocação
            }
        }

        return n; // box
    }

    public static object Bool(bool b) => b ? BoxedTrue : BoxedFalse;
}

public sealed class NumBc(double value) : INodeObj
{
    public object? Eval() => BoxCache.Number(value);
}

public sealed class AddBc(INodeObj left, INodeObj right) : INodeObj
{
    public object? Eval()
    {
        var lv = left.Eval();
        if (lv is SpikeError)
        {
            return lv;
        }

        var rv = right.Eval();
        if (rv is SpikeError)
        {
            return rv;
        }

        return BoxCache.Number(SpikeCoercion.ToNumber(lv) + SpikeCoercion.ToNumber(rv));
    }
}

public sealed class SumBc(INodeObj[] children) : INodeObj
{
    public object? Eval()
    {
        var acc = 0d;
        foreach (var child in children)
        {
            var v = child.Eval();
            if (v is SpikeError)
            {
                return v;
            }

            acc += SpikeCoercion.ToNumber(v);
        }

        return BoxCache.Number(acc);
    }
}

public static class SpikeTreesBoxCache
{
    public static INodeObj SumFold(int n)
    {
        var kids = new INodeObj[n];
        for (var i = 0; i < n; i++)
        {
            kids[i] = new NumBc(i + 1);
        }

        return new SumBc(kids);
    }

    public static INodeObj CumChain(int depth)
    {
        INodeObj node = new NumBc(1);
        for (var i = 2; i <= depth; i++)
        {
            node = new AddBc(node, new NumBc(i));
        }

        return node;
    }

    public static INodeObj Mixed(int n)
    {
        var kids = new INodeObj[n];
        for (var i = 0; i < n; i++)
        {
            kids[i] = i % 20 == 0 ? new TextObj("2") : new NumBc(i + 1);
        }

        return new SumBc(kids);
    }
}
