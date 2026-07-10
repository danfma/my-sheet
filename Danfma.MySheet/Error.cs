using Danfma.MySheet.Expressions;

namespace Danfma.MySheet;

/// <summary>
/// Identidade de um erro do Excel (#VALUE!, #DIV/0!, …) como um value type opaco e alloc-free: embrulha um
/// <c>int</c> código, então cabe inteiro num <see cref="ComputedValue"/> sem alocar no heap. Os erros
/// well-known são instâncias estáticas nomeadas; <see cref="ToString"/> imprime o <see cref="Display"/>.
/// Registro de erros customizados (para custom functions) é uma extensão futura, deliberadamente fora do escopo.
/// </summary>
public readonly struct Error : IEquatable<Error>
{
    // Índice = código. Ordem espelha os valores clássicos do Excel.
    private static readonly string[] Displays =
    [
        "#NULL!", // 0
        "#DIV/0!", // 1
        "#VALUE!", // 2
        "#REF!", // 3
        "#NAME?", // 4
        "#NUM!", // 5
        "#N/A", // 6
    ];

    private readonly int _code;

    private Error(int code) => _code = code;

    public static readonly Error Null = new(0);
    public static readonly Error DivZero = new(1);
    public static readonly Error Value = new(2);
    public static readonly Error Ref = new(3);
    public static readonly Error Name = new(4);
    public static readonly Error Num = new(5);
    public static readonly Error NA = new(6);

    /// <summary>Representação Excel do erro, ex.: <c>"#VALUE!"</c>.</summary>
    public string Display => (uint)_code < (uint)Displays.Length ? Displays[_code] : "#ERR?";

    public override string ToString() => Display;

    public bool Equals(Error other) => _code == other._code;

    public override bool Equals(object? obj) => obj is Error other && Equals(other);

    public override int GetHashCode() => _code;

    public static bool operator ==(Error left, Error right) => left._code == right._code;

    public static bool operator !=(Error left, Error right) => left._code != right._code;

    // --- Pontes internas com o resto do engine (código <-> struct <-> nó de AST ErrorValue) ---

    internal int Code => _code;

    internal static Error FromCode(int code) => new(code);

    /// <summary>Mapeia o código string de um <see cref="ErrorValue"/> (nó de AST) para o <see cref="Error"/>.</summary>
    internal static Error FromDisplay(string display) =>
        display switch
        {
            "#NULL!" => Null,
            "#DIV/0!" => DivZero,
            "#VALUE!" => Value,
            "#REF!" => Ref,
            "#NAME?" => Name,
            "#NUM!" => Num,
            "#N/A" => NA,
            _ => Value,
        };

    /// <summary>Converte para o nó de AST <see cref="ErrorValue"/>, reusando os singletons quando existem.</summary>
    internal ErrorValue ToErrorValue() =>
        _code switch
        {
            1 => ErrorValue.DivByZero,
            2 => ErrorValue.NotValue,
            3 => ErrorValue.Reference,
            4 => ErrorValue.Name,
            5 => ErrorValue.Number,
            6 => ErrorValue.NotAvailable,
            _ => new ErrorValue(Display),
        };
}
