namespace Danfma.MySheet;

/// <summary>
/// Options for <see cref="Workbook.Load(string, WorkbookLoadOptions)"/> and its async counterpart. A file
/// written by any <c>Save</c> overload, with any <see cref="WorkbookSaveOptions"/>, loads identically
/// regardless of these options — <see cref="IoBuffering"/> only changes the read MECHANISM, never the
/// resulting workbook.
/// </summary>
public sealed class WorkbookLoadOptions
{
    /// <summary>
    /// The buffering strategy used to read the file. Defaults to <see cref="WorkbookIoBuffering.Pooled"/> —
    /// see <see cref="WorkbookIoBuffering"/> for the trade-off against <see cref="WorkbookIoBuffering.Pipelines"/>.
    /// </summary>
    public WorkbookIoBuffering IoBuffering { get; init; } = WorkbookIoBuffering.Pooled;
}
