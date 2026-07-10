using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;

namespace Danfma.MySheet.Excel;

/// <summary>
/// Reads <c>xl/sharedStrings.xml</c> forward-only into a presized list, without materializing the
/// OpenXML DOM. Rich-text runs are flattened by concatenating every <c>&lt;t&gt;</c> under the
/// <c>&lt;si&gt;</c> — matching <c>SharedStringItem.InnerText</c>, including phonetic-run
/// (<c>&lt;rPh&gt;</c>) text, so this is a byte-for-byte parity replacement for the DOM read.
/// <c>xml:space="preserve"</c> needs no special handling: <see cref="XmlReader"/> never trims
/// text content (whitespace nodes are only dropped by <c>IgnoreWhitespace</c>, which stays off).
/// </summary>
internal static class SharedStringsStreamReader
{
    public static IReadOnlyList<string> Read(SharedStringTablePart? part)
    {
        if (part is null)
        {
            return [];
        }

        using var source = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = XmlReader.Create(source);

        var strings = new List<string>();
        var builder = new StringBuilder();

        reader.MoveToContent(); // the <sst> root

        // uniqueCount is a capacity HINT only (third-party producers omit or mis-count it); the list
        // still grows past it, and an absurd value is capped instead of trusted.
        if (
            int.TryParse(reader.GetAttribute("uniqueCount"), out var uniqueCount)
            && uniqueCount > 0
        )
        {
            strings.Capacity = Math.Min(uniqueCount, 1 << 20);
        }

        if (reader.IsEmptyElement)
        {
            return strings;
        }

        reader.Read(); // first child of <sst>

        while (!reader.EOF && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == 0))
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si")
            {
                strings.Add(ReadFlattenedText(reader, builder));
            }
            else if (reader.NodeType == XmlNodeType.Element)
            {
                reader.Skip();
            }
            else
            {
                reader.Read();
            }
        }

        return strings;
    }

    /// <summary>
    /// Reads the text container the reader is positioned on (<c>&lt;si&gt;</c> or <c>&lt;is&gt;</c>),
    /// concatenating the content of every <c>&lt;t&gt;</c> at any depth in document order — the
    /// InnerText flattening rule. Consumes through the container's end element and past it. The common
    /// case (a single <c>&lt;t&gt;</c>) never touches the builder.
    /// </summary>
    internal static string ReadFlattenedText(XmlReader reader, StringBuilder builder)
    {
        if (reader.IsEmptyElement)
        {
            reader.Read();

            return string.Empty;
        }

        var containerDepth = reader.Depth;
        string? single = null;
        var multiple = false;

        reader.Read();

        while (
            !reader.EOF
            && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == containerDepth)
        )
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t")
            {
                string text;

                if (reader.IsEmptyElement)
                {
                    text = string.Empty;
                    reader.Read();
                }
                else
                {
                    // Consumes through </t> and lands on the NEXT node; the loop re-inspects the
                    // current position without another Read().
                    text = reader.ReadElementContentAsString();
                }

                if (single is null && !multiple)
                {
                    single = text;
                }
                else
                {
                    if (!multiple)
                    {
                        builder.Clear();
                        builder.Append(single);
                        multiple = true;
                    }

                    builder.Append(text);
                }

                continue;
            }

            if (!reader.Read())
            {
                break;
            }
        }

        reader.Read(); // past the container's end element

        return multiple ? builder.ToString() : single ?? string.Empty;
    }
}
