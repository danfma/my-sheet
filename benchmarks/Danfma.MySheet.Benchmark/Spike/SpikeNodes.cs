using System.Globalization;

namespace Danfma.MySheet.Benchmark.Spike;

// Dois mini-ASTs deliberadamente equivalentes. A ÚNICA diferença de alocação entre eles é o box do
// double que a variante `object?` produz a cada resultado numérico; a variante `CellValue` devolve o
// valor inline. Erros são singletons em cache nas duas variantes (como os ErrorValue reais do MySheet),
// e Text/Reference guardam uma referência já existente — então nem erro nem texto criam diferença de
// alocação, isolando o box numérico como a variável medida.

/// <summary>Erro do spike com instâncias em cache (espelha os ErrorValue singleton do MySheet).</summary>
public sealed class SpikeError
{
    private static readonly SpikeError[] Cache = [new(0), new(1), new(2), new(3), new(4), new(5)];

    private SpikeError(int code) => Code = code;

    public int Code { get; }

    public static SpikeError Of(int code) => Cache[code];
}

// ---------------------------------------------------------------------------
// Variante A — baseline com object? (boxing a cada resultado numérico)
// ---------------------------------------------------------------------------

public interface INodeObj
{
    object? Eval();
}

public sealed class NumObj(double value) : INodeObj
{
    public object? Eval() => value; // box do double
}

public sealed class BoolObj(bool value) : INodeObj
{
    public object? Eval() => value; // box do bool
}

public sealed class TextObj(string value) : INodeObj
{
    public object? Eval() => value; // referência já existente
}

public sealed class ErrObj(int code) : INodeObj
{
    private readonly SpikeError _error = SpikeError.Of(code);

    public object? Eval() => _error; // singleton em cache — sem alocação
}

public sealed class AddObj(INodeObj left, INodeObj right) : INodeObj
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

        return SpikeCoercion.ToNumber(lv) + SpikeCoercion.ToNumber(rv); // box do resultado
    }
}

public sealed class SumObj(INodeObj[] children) : INodeObj
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

        return acc; // box do resultado
    }
}

// ---------------------------------------------------------------------------
// Variante B — experimento com CellValue (sem boxing no caminho numérico)
// ---------------------------------------------------------------------------

public interface INodeCv
{
    CellValue Eval();
}

public sealed class NumCv(double value) : INodeCv
{
    public CellValue Eval() => CellValue.Number(value);
}

public sealed class BoolCv(bool value) : INodeCv
{
    public CellValue Eval() => CellValue.Bool(value);
}

public sealed class TextCv(string value) : INodeCv
{
    public CellValue Eval() => CellValue.Text(value);
}

public sealed class ErrCv(int code) : INodeCv
{
    public CellValue Eval() => CellValue.Error(code);
}

public sealed class AddCv(INodeCv left, INodeCv right) : INodeCv
{
    public CellValue Eval()
    {
        var lv = left.Eval();
        if (lv.IsError)
        {
            return lv;
        }

        var rv = right.Eval();
        if (rv.IsError)
        {
            return rv;
        }

        return CellValue.Number(SpikeCoercion.ToNumber(in lv) + SpikeCoercion.ToNumber(in rv));
    }
}

public sealed class SumCv(INodeCv[] children) : INodeCv
{
    public CellValue Eval()
    {
        var acc = 0d;
        foreach (var child in children)
        {
            var v = child.Eval();
            if (v.IsError)
            {
                return v;
            }

            acc += SpikeCoercion.ToNumber(in v);
        }

        return CellValue.Number(acc);
    }
}

// ---------------------------------------------------------------------------
// Coerção — mesma semântica nas duas variantes (número/bool rápido, texto numérico parseia, blank = 0)
// ---------------------------------------------------------------------------

internal static class SpikeCoercion
{
    public static double ToNumber(object? value) =>
        value switch
        {
            double d => d,
            bool b => b ? 1d : 0d,
            string s => double.Parse(s, CultureInfo.InvariantCulture),
            null => 0d,
            _ => 0d,
        };

    public static double ToNumber(in CellValue value)
    {
        if (value.TryGetNumber(out var n)) // Number | Boolean — rápido, sem unbox
        {
            return n;
        }

        if (value.TryGetText(out var s))
        {
            return double.Parse(s, CultureInfo.InvariantCulture);
        }

        return 0d; // Blank
    }
}

// ---------------------------------------------------------------------------
// Construtores de árvore — as duas variantes recebem as MESMAS formas de carga
// ---------------------------------------------------------------------------

public static class SpikeTrees
{
    /// <summary>Fold largo: um nó que soma N literais (espelha SUM sobre um range). Valor = 1+2+…+N.</summary>
    public static INodeObj SumFoldObj(int n)
    {
        var kids = new INodeObj[n];
        for (var i = 0; i < n; i++)
        {
            kids[i] = new NumObj(i + 1);
        }

        return new SumObj(kids);
    }

    public static INodeCv SumFoldCv(int n)
    {
        var kids = new INodeCv[n];
        for (var i = 0; i < n; i++)
        {
            kids[i] = new NumCv(i + 1);
        }

        return new SumCv(kids);
    }

    /// <summary>Cadeia cumulativa profundidade N: (((1)+2)+3)…+N (espelha coluna B2=B1+A2). Valor = 1+…+N.</summary>
    public static INodeObj CumChainObj(int depth)
    {
        INodeObj node = new NumObj(1);
        for (var i = 2; i <= depth; i++)
        {
            node = new AddObj(node, new NumObj(i));
        }

        return node;
    }

    public static INodeCv CumChainCv(int depth)
    {
        INodeCv node = new NumCv(1);
        for (var i = 2; i <= depth; i++)
        {
            node = new AddCv(node, new NumCv(i));
        }

        return node;
    }

    /// <summary>Fold numérico-dominante com ~5% de texto numérico ("2") para exercitar o ramo de coerção de texto.</summary>
    public static INodeObj MixedObj(int n)
    {
        var kids = new INodeObj[n];
        for (var i = 0; i < n; i++)
        {
            kids[i] = i % 20 == 0 ? new TextObj("2") : new NumObj(i + 1);
        }

        return new SumObj(kids);
    }

    public static INodeCv MixedCv(int n)
    {
        var kids = new INodeCv[n];
        for (var i = 0; i < n; i++)
        {
            kids[i] = i % 20 == 0 ? new TextCv("2") : new NumCv(i + 1);
        }

        return new SumCv(kids);
    }
}
