namespace Danfma.MySheet.Benchmark.Spike;

/// <summary>
/// Discriminador do conteúdo de um <see cref="CellValue"/>. Number/Boolean/Error/Blank vivem
/// inteiramente no campo <c>double</c> + tag (caminho quente sem alocação); Text/Reference guardam
/// uma referência já existente no campo <c>object?</c>.
/// </summary>
public enum CellValueKind : byte
{
    Blank,
    Number,
    Boolean,
    Error,
    Text,
    Reference,
}

/// <summary>
/// Value type opaco que emula uma union sem <c>Unsafe</c>/<c>FieldOffset</c>/<c>IntPtr</c>. Dois campos
/// (um <c>double</c> e um <c>object?</c>) mais uma tag de 1 byte: o GC preciso do CLR proíbe sobrepor um
/// campo gerenciado com um não-gerenciado no mesmo offset, então NÃO se tenta uma union real. Number, Bool,
/// Blank e Error cabem no <c>double</c>; só Text/Reference tocam o campo de referência (sem alocar nada novo).
/// </summary>
public readonly struct CellValue
{
    private readonly double _num; // Number | Boolean(0/1) | Error(código) | Blank(ignorado)
    private readonly object? _ref; // Text(string) | Reference — null nos escalares
    private readonly CellValueKind _kind;

    private CellValue(double num, object? reference, CellValueKind kind)
    {
        _num = num;
        _ref = reference;
        _kind = kind;
    }

    public static readonly CellValue Blank = new(0d, null, CellValueKind.Blank);

    public static CellValue Number(double n) => new(n, null, CellValueKind.Number);

    public static CellValue Bool(bool b) => new(b ? 1d : 0d, null, CellValueKind.Boolean);

    public static CellValue Error(int code) => new(code, null, CellValueKind.Error);

    public static CellValue Text(string s) => new(0d, s, CellValueKind.Text);

    public CellValueKind Kind => _kind;

    public bool IsError => _kind == CellValueKind.Error;

    /// <summary>Caminho quente: lê o número inline, sem unbox e sem checagem de tipo por reflexão.</summary>
    public bool TryGetNumber(out double value)
    {
        value = _num;
        return _kind is CellValueKind.Number or CellValueKind.Boolean;
    }

    public bool TryGetText(out string value)
    {
        if (_kind == CellValueKind.Text)
        {
            value = (string)_ref!;
            return true;
        }

        value = string.Empty;
        return false;
    }

    public bool TryGetError(out int code)
    {
        code = (int)_num;
        return _kind == CellValueKind.Error;
    }

    /// <summary>Ponte para call sites que ainda esperam <c>object?</c> (migração gradual futura).</summary>
    public object? AsObject() =>
        _kind switch
        {
            CellValueKind.Number => _num,
            CellValueKind.Boolean => _num != 0d,
            CellValueKind.Blank => null,
            _ => _ref,
        };
}
