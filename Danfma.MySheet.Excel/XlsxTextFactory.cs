using DocumentFormat.OpenXml;
using XlsxText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace Danfma.MySheet.Excel;

/// <summary>
/// Builds the OpenXML <c>&lt;t&gt;</c> text element for cell strings, tagging it with
/// <c>xml:space="preserve"</c> whenever the value carries leading or trailing whitespace. Without that
/// attribute, spec-compliant readers (Excel, Aspose.Cells) trim the edge whitespace off the cell — only
/// lenient readers like ClosedXML keep it — so a value such as <c>"  foo "</c> would silently lose its
/// spaces on the way into a real spreadsheet.
/// </summary>
internal static class XlsxTextFactory
{
    public static XlsxText Create(string value)
    {
        var text = new XlsxText(value);

        if (value != value.Trim())
        {
            text.Space = SpaceProcessingModeValues.Preserve;
        }

        return text;
    }
}
