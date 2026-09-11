using DocumentFormat.OpenXml.Packaging;
using XlsxSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;

namespace Danfma.MySheet.Excel.Tests;

/// <summary>
/// Package-level probes for assertions that must look at what an .xlsx file physically carries (a
/// <c>&lt;table&gt;</c> part, a <c>&lt;tableParts&gt;</c> element) rather than at what a reader makes of it.
/// </summary>
internal static class XlsxParts
{
    /// <summary>The committed oracle-authored fixture <paramref name="name"/> (see Fixtures/README.md).</summary>
    public static string Fixture(string name) =>
        Path.Combine(AppContext.BaseDirectory, "Fixtures", name + ".xlsx");

    /// <summary>
    /// The worksheet part behind the sheet named <paramref name="sheetName"/>. Always by name, never
    /// <c>WorksheetParts.First()</c>: part enumeration follows relationship order, which is not tab order
    /// (measured: <c>rId4</c> enumerated before <c>rId3</c>).
    /// </summary>
    public static WorksheetPart WorksheetFor(SpreadsheetDocument document, string sheetName)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart
            .Workbook!.Sheets!.Elements<XlsxSheet>()
            .Single(candidate => candidate.Name?.Value == sheetName);

        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
    }

    /// <summary>The raw XML text of a part, for assertions on elements the streaming loader never models.</summary>
    public static string ReadXml(OpenXmlPart part)
    {
        using var reader = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read));

        return reader.ReadToEnd();
    }
}
