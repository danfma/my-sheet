using System.Diagnostics.CodeAnalysis;
using MemoryPack;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// The area of a table a structured reference names: the four Excel specifiers (<c>[#Data]</c> is the
/// default when none is written) plus the two legal specifier PAIRS, <c>[[#Headers],[#Data]]</c> and
/// <c>[[#Data],[#Totals]]</c>. <see cref="Data"/> is 0 so the overwhelmingly common <c>T[Col]</c> form
/// serializes the enum's default byte; the members are append-only because the value is on the wire.
/// </summary>
public enum TableArea : byte
{
    Data = 0,
    All = 1,
    Headers = 2,
    Totals = 3,
    HeadersAndData = 4,
    DataAndTotals = 5,
}

/// <summary>
/// A structured (table) reference — <c>Tabela1[Valor]</c>, <c>Tabela1[#All]</c>,
/// <c>Tabela1[[#Headers],[#Data],[Valor]]</c> — resolved at evaluation time against
/// <see cref="Workbook.Tables"/> to a concrete <see cref="RangeReference"/>, or, for a band with zero rows,
/// to an <see cref="EmptyRangeReference"/>. Contract with the parser:
/// <see cref="TableName"/> and <see cref="ColumnName"/> hold the DECODED payload (the <c>'</c>-prefix escape
/// table <c>'[ '] '# '' '@</c> already applied, exactly as <c>Tokenizer.ReadQuotedName</c> stores decoded
/// text), because the column lookup compares against the raw <c>tableColumn/@name</c> of the xlsx.
/// <c>Area = Data</c> with a non-null <see cref="ColumnName"/> is <c>T[Col]</c>; <c>Area = Data</c> with a
/// null <see cref="ColumnName"/> is <c>T[#Data]</c> (and <c>T[]</c>, which Excel stores as the bare name).
/// Table names are workbook-scoped, like <see cref="Workbook.DefinedNames"/>, so there is no sheet member:
/// the sheet comes from the registered <see cref="Table"/>.
/// <para>
/// <see cref="Expression.IsVolatile"/> is deliberately NOT overridden: a table's target comes from REGISTERED
/// state, whose invalidation is the definitions version bump <c>Workbook.DefineTable</c> performs, not
/// per-pass volatility. Overriding it would make <c>DependencyExtractor</c> mark every structured-reference
/// formula always-dirty and throw away the static range dependency the resolved rectangle provides.
/// </para>
/// </summary>
[MemoryPackable]
public sealed partial record TableReference(string TableName, string? ColumnName, TableArea Area)
    : Reference
{
    /// <summary>
    /// The ONE resolution primitive, so the <c>#NAME?</c>/<c>#REF!</c> mapping lives in exactly one place and
    /// the bounds arithmetic in exactly one other (<see cref="Table.GetRegion"/>, which owns all six areas).
    /// Takes a <see cref="Workbook"/>, not an <see cref="EvaluationContext"/>, because table resolution is
    /// context-free — which is what lets the dependency extractor (which has only a workbook) emit a real
    /// static dependency. It answers one of three things:
    /// <list type="bullet">
    /// <item>a <see cref="RangeReference"/> for an area with rows;</item>
    /// <item>an <see cref="EmptyRangeReference"/> for an area that exists but spans zero rows (sweep item 33 —
    /// a header-only table's <c>[#Data]</c>, measured on the oracle as an EMPTY reference: <c>SUM</c> 0,
    /// <c>ROWS</c> 0, <c>ISREF</c> TRUE);</item>
    /// <item><c>false</c> with the error for an area that does not resolve: an unknown table is
    /// <see cref="Error.Name"/> (Excel resolves a table name in the same name space as a defined name), and a
    /// region that does not exist is <see cref="Error.Ref"/>, Excel's own repair marker (a deleted column's
    /// specifier becomes <c>Table1[#REF!]</c>).</item>
    /// </list>
    /// Every caller states which rectangle it can use (<c>is RangeReference</c> for the cell-streaming fast
    /// paths), so no caller can mistake an empty area for an error or for a real rectangle.
    /// </summary>
    internal bool TryResolve(
        Workbook workbook,
        [NotNullWhen(true)] out Reference? reference,
        out Error error
    )
    {
        reference = null;

        if (!workbook.Tables.TryGetValue(TableName, out var table))
        {
            error = Error.Name;
            return false;
        }

        switch (
            table.GetRegion(
                ColumnName,
                Area,
                out var left,
                out var top,
                out var right,
                out var bottom
            )
        )
        {
            case TableRegionOutcome.Resolved:
                reference = new RangeReference(
                    new CellAddress(left, top).ToId(),
                    new CellAddress(right, bottom).ToId(),
                    table.SheetName
                );
                error = default;
                return true;

            case TableRegionOutcome.Empty:
                // Sweep item 33, reopening Phase 5 ruling R1: the oracle answers an EMPTY reference for a
                // band that spans zero rows (measured on a header-only table, both entry modes: SUM 0,
                // ROWS 0, COLUMNS the band's width, ISREF TRUE, ROW the row after the header). The band's
                // geometry becomes the zero-row rectangle anchored there — never a RangeReference, whose
                // normalized corners would read the header row.
                reference = new EmptyRangeReference(table.SheetName, top, left, right);
                error = default;
                return true;

            default:
                // Absent: an unknown column, or [#Headers]/[#Totals] on a table that has no such row.
                // Measured, both entry modes: SUM(T[#Totals]) over a table with no totals row is #REF! and
                // ISREF is FALSE — on a header-only table too.
                error = Error.Ref;
                return false;
        }
    }

    /// <summary>
    /// The rectangle-only view of <see cref="TryResolve"/>, for the callers that stream a REAL range's cells
    /// (the performance normalizations, the dependency graph, ISFORMULA's top-left cell): <c>true</c> only
    /// when the area resolves to a <see cref="RangeReference"/>. An empty area and an unresolvable table are
    /// both <c>false</c> here, and each caller's own fallback handles them — the empty reference's value
    /// streams nothing, the error value is the node's own.
    /// </summary>
    internal bool TryResolveRectangle(
        Workbook workbook,
        [NotNullWhen(true)] out RangeReference? range
    )
    {
        range = TryResolve(workbook, out var reference, out _) ? reference as RangeReference : null;
        return range is not null;
    }

    public override bool TryResolveReference(EvaluationContext context, out Reference? reference) =>
        TryResolve(context.Workbook, out reference, out _);

    /// <summary>
    /// Two invariants this must never break. (a) It returns the CONCRETE resolved reference, never
    /// <c>ComputedValue.Reference(this)</c>: measured, a node that returned itself made <c>SUM(node)</c>
    /// answer 0, because <c>ComputedValue.EnumerateValues</c>' catch-all <c>case Reference</c> yields the
    /// reference value back as one non-numeric element that the numeric fold silently drops, whereas the
    /// concrete <see cref="RangeReference"/> hits the range arm and expands (and the concrete
    /// <see cref="EmptyRangeReference"/> hits its own arm and yields nothing). (b) It never answers
    /// <c>#VALUE!</c> the way <see cref="RangeReference.Evaluate"/> does: an unresolvable table is the
    /// error VALUE (<c>#NAME?</c> or <c>#REF!</c>) and each consumer does with it what it does with any
    /// error-valued argument — never a throw, never a short-circuit elsewhere.
    /// </summary>
    public override ComputedValue Evaluate(EvaluationContext context) =>
        TryResolve(context.Workbook, out var reference, out var error)
            ? ComputedValue.Reference(reference)
            : ComputedValue.Error(error);
}
