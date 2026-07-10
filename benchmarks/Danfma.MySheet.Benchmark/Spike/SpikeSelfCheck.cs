namespace Danfma.MySheet.Benchmark.Spike;

/// <summary>
/// Sanidade da Fase 1: prova que as variantes <c>object?</c> e <c>CellValue</c> computam os MESMOS valores
/// (número, texto e erro). Garante que o benchmark comparará caminhos equivalentes, não um mais curto.
/// Roda via <c>dotnet run -- --check</c>; lança em caso de divergência.
/// </summary>
public static class SpikeSelfCheck
{
    public static void Run()
    {
        // SumFold(1..100) = 5050 nas duas variantes.
        AssertNumber("SumFold(100) A", SpikeTrees.SumFoldObj(100).Eval(), 5050d);
        AssertNumber("SumFold(100) B", SpikeTrees.SumFoldCv(100).Eval(), 5050d);

        // CumChain(1..100) = 5050 nas três variantes.
        AssertNumber("CumChain(100) A", SpikeTrees.CumChainObj(100).Eval(), 5050d);
        AssertNumber("CumChain(100) B", SpikeTrees.CumChainCv(100).Eval(), 5050d);

        // Variante C (box-cache) equivale a A nos folds/cadeias.
        AssertNumber("SumFold(100) C", SpikeTreesBoxCache.SumFold(100).Eval(), 5050d);
        AssertNumber("CumChain(100) C", SpikeTreesBoxCache.CumChain(100).Eval(), 5050d);
        AssertEqual(
            "Mixed(100) A vs C",
            ToNumber(SpikeTrees.MixedObj(100).Eval()),
            ToNumber(SpikeTreesBoxCache.Mixed(100).Eval())
        );

        // Mixed(100): A e B têm de bater entre si (numérico + texto numérico).
        var mixedA = ToNumber(SpikeTrees.MixedObj(100).Eval());
        var mixedB = ToNumber(SpikeTrees.MixedCv(100).Eval());
        AssertEqual("Mixed(100) A vs B", mixedA, mixedB);

        // Coerção de texto: "3" + 4 = 7 nas duas variantes.
        AssertNumber("Text add A", new AddObj(new TextObj("3"), new NumObj(4)).Eval(), 7d);
        AssertNumber("Text add B", new AddCv(new TextCv("3"), new NumCv(4)).Eval(), 7d);

        // Propagação de erro: [1, #erro(3), 2] curto-circuita para o erro de código 3 nas duas variantes.
        AssertError(
            "Error propagation A",
            new SumObj([new NumObj(1), new ErrObj(3), new NumObj(2)]).Eval(),
            3
        );
        AssertError(
            "Error propagation B",
            new SumCv([new NumCv(1), new ErrCv(3), new NumCv(2)]).Eval(),
            3
        );

        // Sonda de profundidade: CumChain(3000) recursiona 3000 níveis. Se sobrevive aqui (stack default),
        // o worker do BenchmarkDotNet (mesma stack default) também sobrevive → sem stack overflow no run.
        const int deep = 3000;
        var deepExpected = deep * (deep + 1) / 2d; // 4.501.500
        AssertNumber("CumChain(3000) A", SpikeTrees.CumChainObj(deep).Eval(), deepExpected);
        AssertNumber("CumChain(3000) B", SpikeTrees.CumChainCv(deep).Eval(), deepExpected);
        AssertNumber("CumChain(3000) C", SpikeTreesBoxCache.CumChain(deep).Eval(), deepExpected);

        // Grafo memoizado (ranges + dados mistos + caminhos cruzados): os dois motores de cache
        // (object? vs CellValue) têm de produzir o MESMO checksum de extração.
        var (cells, ids) = GraphBuilder.Build(5000);
        var objChecksum = new ObjEngine(cells, ids).ExtractAll();
        var cvChecksum = new CvEngine(cells, ids).ExtractAll();
        if (double.IsNaN(objChecksum) || double.IsInfinity(objChecksum))
        {
            throw new InvalidOperationException($"Checksum do grafo degenerou: {objChecksum}.");
        }

        AssertEqual("Graph ExtractAll object? vs CellValue", objChecksum, cvChecksum);

        Console.WriteLine(
            "Spike self-check OK — object?, box-cache e CellValue equivalentes (inclui sonda de profundidade 3000 e grafo memoizado 5000)."
        );
    }

    private static double ToNumber(object? value) =>
        value switch
        {
            double d => d,
            _ => throw new InvalidOperationException(
                $"Esperado número, veio {value?.GetType().Name ?? "null"}."
            ),
        };

    private static double ToNumber(CellValue value) =>
        value.TryGetNumber(out var n)
            ? n
            : throw new InvalidOperationException($"Esperado número, veio {value.Kind}.");

    private static void AssertNumber(string label, object? actual, double expected) =>
        AssertEqual(label, ToNumber(actual), expected);

    private static void AssertNumber(string label, CellValue actual, double expected) =>
        AssertEqual(label, ToNumber(actual), expected);

    private static void AssertEqual(string label, double actual, double expected)
    {
        if (Math.Abs(actual - expected) > 1e-9)
        {
            throw new InvalidOperationException($"{label}: esperado {expected}, veio {actual}.");
        }
    }

    private static void AssertError(string label, object? actual, int expectedCode)
    {
        if (actual is not SpikeError { } error || error.Code != expectedCode)
        {
            throw new InvalidOperationException(
                $"{label}: esperado erro código {expectedCode}, veio {actual}."
            );
        }
    }

    private static void AssertError(string label, CellValue actual, int expectedCode)
    {
        if (!actual.TryGetError(out var code) || code != expectedCode)
        {
            throw new InvalidOperationException(
                $"{label}: esperado erro código {expectedCode}, veio {actual.Kind}."
            );
        }
    }
}
