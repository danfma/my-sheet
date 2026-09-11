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
/// <see cref="Workbook.Tables"/> to a concrete <see cref="RangeReference"/>. Contract with the parser:
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
    /// The ONE resolution primitive, so the <c>#NAME?</c>/<c>#REF!</c> mapping and the bounds arithmetic
    /// exist exactly once. Takes a <see cref="Workbook"/>, not an <see cref="EvaluationContext"/>, because
    /// table resolution is context-free — which is what lets the dependency extractor (which has only a
    /// workbook) emit a real static dependency. An unknown table is <see cref="Error.Name"/> (Excel resolves
    /// a table name in the same name space as a defined name); an unknown column, or a
    /// <see cref="TableArea.Data"/> area over a table with no data rows, is <see cref="Error.Ref"/> (Excel's
    /// own repair rewrites a deleted column's specifier to <c>Table1[#REF!]</c>; the header-only case is a
    /// recorded divergence, see the tests).
    /// </summary>
    internal bool TryResolveRange(Workbook workbook, out RangeReference? range, out Error error)
    {
        range = null;

        if (!workbook.Tables.TryGetValue(TableName, out var table))
        {
            error = Error.Name;
            return false;
        }

        // Phase 3 shipped the [#Data] geometry only (Table.TryGetColumnRange and the derived data rows).
        // The other five areas answer #REF! until Phase 5 T1 lands Table.TryGetRegion, which owns their
        // measured geometry (the pairs SHRINK to the data body when the row they name is absent; only the
        // singletons error) and routes this method through it.
        if (Area != TableArea.Data)
        {
            error = Error.Ref;
            return false;
        }

        if (ColumnName is null)
        {
            // T[#Data] / T[]: the whole data body. Same empty-table rule as TryGetColumnRange.
            if (table.DataRowCount == 0)
            {
                error = Error.Ref;
                return false;
            }

            range = Rectangle(
                table,
                table.FirstColumn,
                table.LastColumn,
                table.FirstDataRow,
                table.LastDataRow
            );
            error = default;
            return true;
        }

        if (!table.TryGetColumnRange(ColumnName, out var column, out var top, out var bottom))
        {
            error = Error.Ref;
            return false;
        }

        range = Rectangle(table, column, column, top, bottom);
        error = default;
        return true;
    }

    // The same construction DynamicRange uses for its resolved rectangle.
    private static RangeReference Rectangle(
        Table table,
        int left,
        int right,
        int top,
        int bottom
    ) =>
        new(
            new CellAddress(left, top).ToId(),
            new CellAddress(right, bottom).ToId(),
            table.SheetName
        );

    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        var ok = TryResolveRange(context.Workbook, out var range, out _);
        reference = range;
        return ok;
    }

    /// <summary>
    /// Two invariants this must never break. (a) It returns the CONCRETE resolved range, never
    /// <c>ComputedValue.Reference(this)</c>: measured, a node that returned itself made <c>SUM(node)</c>
    /// answer 0, because <c>ComputedValue.EnumerateValues</c>' catch-all <c>case Reference</c> yields the
    /// reference value back as one non-numeric element that the numeric fold silently drops, whereas the
    /// concrete <see cref="RangeReference"/> hits the range arm and expands. (b) It never answers
    /// <c>#VALUE!</c> the way <see cref="RangeReference.Evaluate"/> does: an unresolvable table is the
    /// error VALUE (<c>#NAME?</c> or <c>#REF!</c>) and each consumer does with it what it does with any
    /// error-valued argument — never a throw, never a short-circuit elsewhere.
    /// </summary>
    public override ComputedValue Evaluate(EvaluationContext context) =>
        TryResolveRange(context.Workbook, out var range, out var error)
            ? ComputedValue.Reference(range!)
            : ComputedValue.Error(error);
}
