namespace Danfma.MySheet.Benchmark.Spike;

// Cenário fiel ao uso real do MySheet: um grafo de dependências entre células, memoizado, extraído em lote.
// Modela o que o spike anterior NÃO cobria — ranges, dados mistos, caminhos cruzados/compartilhados
// (A←B←C, D←(B,E)) e, sobretudo, o CACHE. Compara duas estratégias de cache:
//   • ObjEngine  → Dictionary<string, object?>   (cada valor numérico vira um box de vida longa)
//   • CvEngine   → Dictionary<string, CellValue> (o struct fica inline na entrada; zero box por valor)
// A extração em ordem topológica (índice crescente; deps sempre em índices anteriores) mantém a recursão
// rasa — reflete o padrão de extração em lote e evita stack overflow.

public enum CellOp : byte
{
    Literal, // número literal (folha)
    LiteralText, // texto numérico "2" (dados mistos)
    Add, // = DepA + DepB (caminhos cruzados/compartilhados)
    SumRange, // = SUM(RangeStart..RangeEnd) (range sobre células cacheadas)
}

public readonly struct CellSpec
{
    public CellOp Op { get; init; }
    public double Literal { get; init; }
    public int DepA { get; init; }
    public int DepB { get; init; }
    public int RangeStart { get; init; }
    public int RangeEnd { get; init; }
}

public static class GraphBuilder
{
    private const int BaseLeaves = 64; // células 0..63 = folhas, viram "hubs" compartilhados por muitas outras

    /// <summary>
    /// Constrói um grafo determinístico de N células: 64 folhas (algumas texto numérico), depois células
    /// que dependem do predecessor imediato E de um hub compartilhado (0..63) — criando D e E que dependem
    /// do mesmo B — mais ~3% de células SUM(range) sobre uma janela de células cacheadas.
    /// </summary>
    public static (CellSpec[] Cells, string[] Ids) Build(int n)
    {
        var cells = new CellSpec[n];
        var ids = new string[n];

        for (var i = 0; i < n; i++)
        {
            ids[i] = "C" + i;

            if (i < BaseLeaves)
            {
                cells[i] =
                    i % 16 == 0
                        ? new CellSpec { Op = CellOp.LiteralText } // dados mistos: texto numérico
                        : new CellSpec { Op = CellOp.Literal, Literal = (i % 10) + 1 };
            }
            else if (i % 32 == 0)
            {
                cells[i] = new CellSpec
                {
                    Op = CellOp.SumRange,
                    RangeStart = i - 16,
                    RangeEnd = i - 1,
                };
            }
            else
            {
                // Caminho cruzado: predecessor imediato (cadeia) + hub compartilhado (fan-in).
                cells[i] = new CellSpec
                {
                    Op = CellOp.Add,
                    DepA = i - 1,
                    DepB = i % BaseLeaves,
                };
            }
        }

        return (cells, ids);
    }
}

/// <summary>Motor com cache <c>object?</c> — cada valor numérico cacheado é um box de vida longa.</summary>
public sealed class ObjEngine
{
    private readonly CellSpec[] _cells;
    private readonly string[] _ids;
    private readonly Dictionary<string, object?> _cache;

    public ObjEngine(CellSpec[] cells, string[] ids)
    {
        _cells = cells;
        _ids = ids;
        _cache = new Dictionary<string, object?>(cells.Length); // pré-dimensionado
    }

    public object? Get(int i)
    {
        var id = _ids[i];
        if (_cache.TryGetValue(id, out var cached))
        {
            return cached;
        }

        ref readonly var spec = ref _cells[i];
        object? computed = spec.Op switch
        {
            CellOp.Literal => spec.Literal, // box
            CellOp.LiteralText => "2", // referência (sem box)
            CellOp.Add => SpikeCoercion.ToNumber(Get(spec.DepA))
                + SpikeCoercion.ToNumber(Get(spec.DepB)), // box
            CellOp.SumRange => SumRange(in spec), // box
            _ => null,
        };

        _cache[id] = computed; // box guardado no cache (vida longa)
        return computed;
    }

    private double SumRange(in CellSpec spec)
    {
        var acc = 0d;
        for (var j = spec.RangeStart; j <= spec.RangeEnd; j++)
        {
            acc += SpikeCoercion.ToNumber(Get(j));
        }

        return acc;
    }

    /// <summary>Extrai todas as células (cada uma computa 1×), devolvendo um checksum. Reusa o cache (Clear).</summary>
    public double ExtractAll()
    {
        _cache.Clear();
        var checksum = 0d;
        for (var i = 0; i < _cells.Length; i++)
        {
            checksum += SpikeCoercion.ToNumber(Get(i));
        }

        return checksum;
    }
}

/// <summary>Motor com cache <c>CellValue</c> — o struct fica inline na entrada; zero box por valor.</summary>
public sealed class CvEngine
{
    private readonly CellSpec[] _cells;
    private readonly string[] _ids;
    private readonly Dictionary<string, CellValue> _cache;

    public CvEngine(CellSpec[] cells, string[] ids)
    {
        _cells = cells;
        _ids = ids;
        _cache = new Dictionary<string, CellValue>(cells.Length);
    }

    public CellValue Get(int i)
    {
        var id = _ids[i];
        if (_cache.TryGetValue(id, out var cached))
        {
            return cached;
        }

        ref readonly var spec = ref _cells[i];
        CellValue computed = spec.Op switch
        {
            CellOp.Literal => CellValue.Number(spec.Literal),
            CellOp.LiteralText => CellValue.Text("2"),
            CellOp.Add => CellValue.Number(
                SpikeCoercion.ToNumber(Get(spec.DepA)) + SpikeCoercion.ToNumber(Get(spec.DepB))
            ),
            CellOp.SumRange => CellValue.Number(SumRange(in spec)),
            _ => CellValue.Blank,
        };

        _cache[id] = computed; // struct inline — sem box
        return computed;
    }

    private double SumRange(in CellSpec spec)
    {
        var acc = 0d;
        for (var j = spec.RangeStart; j <= spec.RangeEnd; j++)
        {
            acc += SpikeCoercion.ToNumber(Get(j));
        }

        return acc;
    }

    public double ExtractAll()
    {
        _cache.Clear();
        var checksum = 0d;
        for (var i = 0; i < _cells.Length; i++)
        {
            checksum += SpikeCoercion.ToNumber(Get(i));
        }

        return checksum;
    }
}
