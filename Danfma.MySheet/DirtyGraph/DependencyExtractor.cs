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
///
/// <para><b>Shared-formula delta (produção, Fase 3).</b> Um <see cref="SharedFormulaSlave"/> NÃO é
/// always-dirty: o walker entra na sua árvore ancorada compartilhada (<see cref="SharedFormulaSlave.Master"/>)
/// carregando o (DeltaRow, DeltaColumn) DA PRÓPRIA escrava, e um <see cref="AnchoredCellReference"/>/
/// <see cref="AnchoredRangeReference"/> encontrado ali dentro aplica esse delta aos componentes RELATIVOS
/// (um componente $-ancorado não se move) para produzir exatamente o mesmo <see cref="CellDep"/>/
/// <see cref="RangeDep"/> que a expansão legada por-célula (<c>ExpressionParser.ParseSharedFormulaBody</c>)
/// produziria — ver <c>SharedFormulaDependencyParityTests</c> para a prova.</para>
/// </summary>
internal static class DependencyExtractor
{
    /// <summary>Extrai as dependências de <paramref name="expression"/>. <paramref name="workbook"/> (opcional)
    /// é usado só para resolver <see cref="NameReference"/> via <see cref="Workbook.DefinedNames"/>; sem ele,
    /// um nome vira always-dirty.</summary>
    public static DependencyScan Extract(Expression expression, Workbook? workbook = null)
    {
        var scan = new DependencyScan();
        Visit(expression, scan, workbook, resolving: null, deltaRow: 0, deltaColumn: 0);
        return scan;
    }

    private static void Visit(
        Expression e,
        DependencyScan scan,
        Workbook? wb,
        HashSet<string>? resolving,
        int deltaRow,
        int deltaColumn
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
                    Visit(area, scan, wb, resolving, deltaRow, deltaColumn);
                }
                return;

            case DynamicRange dynamicRange:
                // Endpoints reference-returning (INDEX(...):A5) — alvo não-estático.
                scan.AlwaysDirty = true;
                Visit(dynamicRange.Start, scan, wb, resolving, deltaRow, deltaColumn);
                Visit(dynamicRange.End, scan, wb, resolving, deltaRow, deltaColumn);
                return;

            case Offset offset:
                // A célula/range alvo é computada de rows/cols/height/width — não-enumerável.
                scan.AlwaysDirty = true;
                foreach (var argument in offset.Arguments)
                {
                    Visit(argument, scan, wb, resolving, deltaRow, deltaColumn);
                }
                return;

            case Indirect indirect:
                // Referência nomeada por texto computado — não-enumerável (e também volátil). Tratada aqui
                // explicitamente por clareza; o VisitArguments genérico ainda blinda qualquer nó não-mapeado.
                scan.AlwaysDirty = true;
                foreach (var argument in indirect.Arguments)
                {
                    Visit(argument, scan, wb, resolving, deltaRow, deltaColumn);
                }
                return;

            case NameReference name:
                ResolveName(name, scan, wb, resolving, deltaRow, deltaColumn);
                return;

            // Shared-formula delta (production): a SharedFormulaSlave contributes exactly the dependencies its
            // shared AnchoredTree would have if independently expanded with ITS OWN (DeltaRow, DeltaColumn) —
            // walk the master with the slave's delta pushed, mirroring how SharedFormulaSlave.Evaluate pushes
            // the same delta into the EvaluationContext (WithDelta) for the runtime path.
            case SharedFormulaSlave slave:
                Visit(slave.Master, scan, wb, resolving, slave.DeltaRow, slave.DeltaColumn);
                return;

            // Only ever reached while walking INSIDE a SharedFormulaSlave.Master tree (see that type's doc
            // comment: the Parser's anchored mode never produces these outside a shared-formula master), so
            // deltaRow/deltaColumn here are always the enclosing slave's own delta. Effective(...) applies the
            // delta to the relative components only (an absolute $-anchored component stays put) — the exact
            // same address a legacy per-slave expansion (ExpressionParser.ParseSharedFormulaBody) would have
            // produced as a plain CellReference.
            case AnchoredCellReference anchoredCell:
                var (effColumn, effRow) = anchoredCell.Effective(deltaRow, deltaColumn);
                scan.Cells.Add(new CellDep(anchoredCell.SheetName, effColumn, effRow));
                return;

            // Same idea as AnchoredCellReference, for a bounded range endpoint pair — EffectiveEndpoints applies
            // the delta per-endpoint (each axis independently, per its own $-anchor flag) and the RangeDep
            // normalizes the corners exactly like the plain RangeReference case above.
            case AnchoredRangeReference anchoredRange:
                var (startColumn, startRow, endColumn, endRow) = anchoredRange.EffectiveEndpoints(
                    deltaRow,
                    deltaColumn
                );
                scan.Ranges.Add(
                    new RangeDep(
                        anchoredRange.SheetName,
                        Math.Min(startColumn, endColumn),
                        Math.Max(startColumn, endColumn),
                        Math.Min(startRow, endRow),
                        Math.Max(startRow, endRow)
                    )
                );
                return;

            case FunctionCall custom:
                // Função host: comportamento desconhecido → conservador.
                scan.AlwaysDirty = true;
                foreach (var argument in custom.Arguments)
                {
                    Visit(argument, scan, wb, resolving, deltaRow, deltaColumn);
                }
                return;

            case BinaryOperation binary:
                Visit(binary.Left, scan, wb, resolving, deltaRow, deltaColumn);
                Visit(binary.Right, scan, wb, resolving, deltaRow, deltaColumn);
                return;

            case UnaryOperation unary:
                Visit(unary.Operand, scan, wb, resolving, deltaRow, deltaColumn);
                return;

            case ValueExpression:
                return; // literais (número, texto, bool, blank, erro) não têm deps

            case Function function:
                // INDIRECT e os voláteis (NOW/TODAY/RAND/RANDBETWEEN) são always-dirty via IsVolatile.
                if (function.IsVolatile)
                {
                    scan.AlwaysDirty = true;
                }
                VisitArguments(function, scan, wb, resolving, deltaRow, deltaColumn);
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
        HashSet<string>? resolving,
        int deltaRow,
        int deltaColumn
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
            Visit(argument, scan, wb, resolving, deltaRow, deltaColumn);
        }
    }

    private static void ResolveName(
        NameReference name,
        DependencyScan scan,
        Workbook? wb,
        HashSet<string>? resolving,
        int deltaRow,
        int deltaColumn
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
            // A defined name's own definition is parsed in ORDINARY (non-anchored) mode and is position-
            // independent (AnchoredFormulaSupport treats NameReference as safe to leave un-anchored) — it
            // never contains an AnchoredCellReference/AnchoredRangeReference, so the delta threaded here is
            // inert in practice. Passed through anyway (rather than reset to 0) to mirror the runtime's own
            // NamedReferences.EvaluateDefinition, which evaluates the definition against the ambient context
            // unchanged.
            Visit(definition, scan, wb, resolving, deltaRow, deltaColumn);
        }
        finally
        {
            resolving.Remove(name.Name);
        }
    }
}
