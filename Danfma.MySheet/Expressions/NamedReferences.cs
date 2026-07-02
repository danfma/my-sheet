using System.Diagnostics.CodeAnalysis;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// Resolution of workbook-level defined names (<see cref="Workbook.DefinedNames"/>). A name maps to any
/// <see cref="Expression"/>; this helper both evaluates that expression to a <see cref="ComputedValue"/>
/// (for <see cref="NameReference"/>) and unwraps it to a syntactic <see cref="Reference"/> (for functions
/// that require a reference node in an argument position, like VLOOKUP's table). Both paths share a
/// thread-local name→name cycle guard, mirroring <c>Workbook._evaluating</c> for cells: a cycle degrades
/// to <c>#REF!</c> instead of overflowing the stack.
/// </summary>
internal static class NamedReferences
{
    // Names currently being resolved on the calling thread, to break name→name cycles. Thread-local so
    // concurrent resolution of the same name on different threads is not a false cycle.
    [ThreadStatic]
    private static HashSet<string>? _resolving;

    /// <summary>
    /// Evaluates a defined name's expression. A range/union stays a reference value (so range-aware
    /// consumers like SUM expand it, exactly as OFFSET's multi-cell result does); anything else — a single
    /// cell, a constant, another name, a formula — evaluates to its scalar value. Guarded against name
    /// cycles; a cycle yields <c>#REF!</c>.
    /// </summary>
    public static ComputedValue EvaluateDefinition(
        Expression definition,
        EvaluationContext context,
        string name
    )
    {
        var guard = _resolving ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (!guard.Add(name))
        {
            return ComputedValue.Error(Error.Ref);
        }

        try
        {
            return definition is RangeReference or UnionReference
                ? ComputedValue.Reference((Reference)definition)
                : definition.Evaluate(context);
        }
        finally
        {
            guard.Remove(name);
        }
    }

    /// <summary>
    /// Unwraps an argument to a syntactic <see cref="Reference"/> node for the functions that need one
    /// (VLOOKUP/HLOOKUP table, INDEX, OFFSET base, ROWS, COLUMNS, AREAS, ISREF). A reference node resolves
    /// to itself; a <see cref="NameReference"/> resolves through the LET scope first (so a LET binding
    /// shadows a defined name) and then the workbook's defined names, recursively (name→name), with the
    /// same cycle guard. Anything else — a constant, an unbound name, a cycle — is not a reference.
    /// </summary>
    public static bool TryResolveReference(
        Expression expression,
        EvaluationContext context,
        [NotNullWhen(true)] out Reference? reference
    )
    {
        if (expression is Reference direct)
        {
            reference = direct;
            return true;
        }

        if (expression is not NameReference name)
        {
            reference = null;
            return false;
        }

        // A LET binding shadows a defined name completely: if the name is bound in scope, only a
        // reference-kind value (e.g. a range captured by LET) counts; a scalar binding is not a reference.
        if (context.TryGetName(name.Name, out var bound))
        {
            return bound.TryGetReference(out reference);
        }

        if (!context.Workbook.DefinedNames.TryGetValue(name.Name, out var definition))
        {
            reference = null;
            return false;
        }

        var guard = _resolving ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (!guard.Add(name.Name))
        {
            reference = null;
            return false;
        }

        try
        {
            return TryResolveReference(definition, context, out reference);
        }
        finally
        {
            guard.Remove(name.Name);
        }
    }

    /// <summary>
    /// Validates a defined name (case-insensitive, workbook-level). Excel-style rules: it must start with a
    /// letter or underscore, contain only letters, digits, '.' or '_', and must not collide with a cell
    /// reference (e.g. <c>A1</c>) or a boolean literal. Throws <see cref="ArgumentException"/> otherwise.
    /// </summary>
    public static void ValidateName(string name)
    {
        ArgumentNullException.ThrowIfNull(name);

        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("A defined name cannot be empty.", nameof(name));
        }

        if (!IsValidName(name))
        {
            throw new ArgumentException(
                $"'{name}' is not a valid defined name: it must start with a letter or underscore, contain "
                    + "only letters, digits, '.' or '_', and must not look like a cell reference (e.g. \"A1\") "
                    + "or a boolean literal.",
                nameof(name)
            );
        }
    }

    private static bool IsValidName(string name)
    {
        if (!char.IsLetter(name[0]) && name[0] != '_')
        {
            return false;
        }

        foreach (var c in name)
        {
            if (!char.IsLetterOrDigit(c) && c is not ('_' or '.'))
            {
                return false;
            }
        }

        // A1-style cell references are reserved (Parser.IsCellReference is the single source of truth), as
        // are the boolean literals the tokenizer would read as BooleanValue instead of a NameReference.
        return !Parser.IsCellReference(name)
            && !string.Equals(name, "TRUE", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(name, "FALSE", StringComparison.OrdinalIgnoreCase);
    }
}
