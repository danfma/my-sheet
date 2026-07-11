using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Lookup;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.DirtyGraph;

/// <summary>Uma dependência de célula única (endereço numérico 1-based, por sheet).</summary>
internal readonly record struct CellDep(string Sheet, int Column, int Row);

/// <summary>
/// Uma dependência de range. Limites <c>null</c> = eixo aberto naquele lado (coluna/linha inteira), igual ao
/// <see cref="OpenRangeReference"/>; um <see cref="RangeReference"/> fechado tem os quatro limites setados.
/// </summary>
internal readonly record struct RangeDep(
    string Sheet,
    int? ColMin,
    int? ColMax,
    int? RowMin,
    int? RowMax
);

/// <summary>
/// O resultado da extração de dependências de UMA fórmula: as células e ranges que ela lê estaticamente, e
/// se ela é <see cref="AlwaysDirty"/> — quando contém um nó cuja dependência NÃO é enumerável estaticamente
/// (OFFSET/INDIRECT/DynamicRange, um volátil, um nome irresolúvel, ou uma função custom de comportamento
/// desconhecido). Uma fórmula always-dirty recomputa em toda passada (reusa o modelo de taint volátil).
/// </summary>
internal sealed class DependencyScan
{
    public HashSet<CellDep> Cells { get; } = [];
    public List<RangeDep> Ranges { get; } = [];
    public bool AlwaysDirty { get; internal set; }

    public bool IsEmpty => Cells.Count == 0 && Ranges.Count == 0 && !AlwaysDirty;
}

/// <summary>
/// Extrai as dependências FORWARD de uma <see cref="Expression"/> (o que a fórmula lê), percorrendo o AST.
///
/// <para><b>Super-aproximação é segura por design.</b> Para a marcação dirty, incluir uma dependência a mais
/// (falso-dirty: recomputar algo que não mudou) é inofensivo; PERDER uma dependência (dirty-perdido) deixa um
/// valor stale — um bug de corretude. Então: coletamos as refs de TODOS os filhos (ambos os ramos de um IF,
/// todos os args de CHOOSE, o range inteiro de um INDEX), e só marcamos <see cref="DependencyScan.AlwaysDirty"/>
/// quando o conjunto de dependências genuinamente não pode ser enumerado.</para>
///
/// <para>Estático (deps enumeráveis): <see cref="CellReference"/>, <see cref="RangeReference"/>,
/// <see cref="OpenRangeReference"/>, <see cref="UnionReference"/>, operadores, e os args de qualquer função
/// built-in (incl. INDEX/CHOOSE/VLOOKUP/MATCH — a dependência é o range inteiro que eles varrem).</para>
///
/// <para>Always-dirty (deps não-enumeráveis): OFFSET e INDIRECT (célula-alvo computada), DynamicRange
/// (endpoints reference-returning), os voláteis NOW/TODAY/RAND/RANDBETWEEN (<see cref="Expression.IsVolatile"/>),
/// um <see cref="FunctionCall"/> custom (comportamento host desconhecido), e um <see cref="NameReference"/>
/// que não resolve.</para>
/// </summary>
internal static class DependencyExtractor
{
    /// <summary>Extrai as dependências de <paramref name="expression"/>. <paramref name="workbook"/> (opcional)
    /// é usado só para resolver <see cref="NameReference"/> via <see cref="Workbook.DefinedNames"/>; sem ele,
    /// um nome vira always-dirty.</summary>
    public static DependencyScan Extract(Expression expression, Workbook? workbook = null)
    {
        var scan = new DependencyScan();
        Visit(expression, scan, workbook, resolving: null);
        return scan;
    }

    private static void Visit(
        Expression e,
        DependencyScan scan,
        Workbook? wb,
        HashSet<string>? resolving
    )
    {
        switch (e)
        {
            case CellReference cell:
                var a = CellAddress.Parse(cell.Id);
                scan.Cells.Add(new CellDep(cell.SheetName, a.Column, a.Row));
                return;

            case RangeReference range:
                var s = CellAddress.Parse(range.StartId);
                var en = CellAddress.Parse(range.EndId);
                scan.Ranges.Add(
                    new RangeDep(
                        range.SheetName,
                        Math.Min(s.Column, en.Column),
                        Math.Max(s.Column, en.Column),
                        Math.Min(s.Row, en.Row),
                        Math.Max(s.Row, en.Row)
                    )
                );
                return;

            case OpenRangeReference open:
                scan.Ranges.Add(
                    new RangeDep(open.SheetName, open.ColMin, open.ColMax, open.RowMin, open.RowMax)
                );
                return;

            case UnionReference union:
                foreach (var area in union.Areas)
                {
                    Visit(area, scan, wb, resolving);
                }
                return;

            case DynamicRange dynamicRange:
                // Endpoints reference-returning (INDEX(...):A5) — alvo não-estático.
                scan.AlwaysDirty = true;
                Visit(dynamicRange.Start, scan, wb, resolving);
                Visit(dynamicRange.End, scan, wb, resolving);
                return;

            case Offset offset:
                // A célula/range alvo é computada de rows/cols/height/width — não-enumerável.
                scan.AlwaysDirty = true;
                foreach (var argument in offset.Arguments)
                {
                    Visit(argument, scan, wb, resolving);
                }
                return;

            case Indirect indirect:
                // Referência nomeada por texto computado — não-enumerável (e também volátil). Tratada aqui
                // explicitamente por clareza; o VisitArguments genérico ainda blinda qualquer nó não-mapeado.
                scan.AlwaysDirty = true;
                foreach (var argument in indirect.Arguments)
                {
                    Visit(argument, scan, wb, resolving);
                }
                return;

            case NameReference name:
                ResolveName(name, scan, wb, resolving);
                return;

            // G3 spike (node-delta shared formulas) — PRODUCTION PENDENCY, not wired for the spike: a
            // SharedFormulaSlave's effective cell/range deps depend on ITS OWN (DeltaRow, DeltaColumn), which
            // this static, delta-oblivious walker does not thread through (unlike FormulaWriter/
            // NumericAggregation/ReferenceGuard, which do — see those files). Conservatively AlwaysDirty
            // instead of silently returning an EMPTY dependency set (the unsafe "lost dependency" failure
            // mode this file's own doc comment calls out) — no test in this repo currently loads a shared
            // formula from .xlsx through DirtyGraph/RecalculationEngine (verified: no "sharedformula" hit
            // under tests/Danfma.MySheet.Tests), so under-implementing this is an explicit, documented
            // tradeoff of the spike, not a silent gap. Wiring full support means threading (deltaRow,
            // deltaColumn) through this whole Visit chain and adding effective-address cases for
            // AnchoredCellReference/AnchoredRangeReference, mirroring the three files above.
            case SharedFormulaSlave
            or AnchoredCellReference
            or AnchoredRangeReference:
                scan.AlwaysDirty = true;
                return;

            case FunctionCall custom:
                // Função host: comportamento desconhecido → conservador.
                scan.AlwaysDirty = true;
                foreach (var argument in custom.Arguments)
                {
                    Visit(argument, scan, wb, resolving);
                }
                return;

            case BinaryOperation binary:
                Visit(binary.Left, scan, wb, resolving);
                Visit(binary.Right, scan, wb, resolving);
                return;

            case UnaryOperation unary:
                Visit(unary.Operand, scan, wb, resolving);
                return;

            case ValueExpression:
                return; // literais (número, texto, bool, blank, erro) não têm deps

            case Function function:
                // INDIRECT e os voláteis (NOW/TODAY/RAND/RANDBETWEEN) são always-dirty via IsVolatile.
                if (function.IsVolatile)
                {
                    scan.AlwaysDirty = true;
                }
                VisitArguments(function, scan, wb, resolving);
                return;

            default:
                return;
        }
    }

    // Argumentos de qualquer função built-in, via o mapa central node→(nome, args) do FormulaWriter. Se um nó
    // não estiver mapeado lá (gap da lib), não dá para enumerar seus args → conservador: marca always-dirty.
    private static void VisitArguments(
        Function function,
        DependencyScan scan,
        Workbook? wb,
        HashSet<string>? resolving
    )
    {
        Expression[] arguments;
        try
        {
            arguments = FormulaWriter.Call(function).Arguments;
        }
        catch (NotSupportedException)
        {
            scan.AlwaysDirty = true;
            return;
        }

        foreach (var argument in arguments)
        {
            Visit(argument, scan, wb, resolving);
        }
    }

    private static void ResolveName(
        NameReference name,
        DependencyScan scan,
        Workbook? wb,
        HashSet<string>? resolving
    )
    {
        if (wb is null || !wb.DefinedNames.TryGetValue(name.Name, out var definition))
        {
            scan.AlwaysDirty = true; // nome irresolúvel (ou sem workbook) → conservador
            return;
        }

        resolving ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!resolving.Add(name.Name))
        {
            scan.AlwaysDirty = true; // ciclo nome→nome
            return;
        }

        try
        {
            Visit(definition, scan, wb, resolving);
        }
        finally
        {
            resolving.Remove(name.Name);
        }
    }
}
