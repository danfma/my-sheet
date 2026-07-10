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

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                continue;
            }

            if (reader.LocalName == "sst")
            {
                // uniqueCount is a capacity HINT only (third-party producers omit or mis-count it);
                // the list still grows past it, and an absurd value is capped instead of trusted.
                if (
                    int.TryParse(reader.GetAttribute("uniqueCount"), out var uniqueCount)
                    && uniqueCount > 0
                )
                {
                    strings.Capacity = Math.Min(uniqueCount, 1 << 20);
                }

                continue;
            }

            if (reader.LocalName == "si")
            {
                strings.Add(ReadItem(reader, builder));
            }
        }

        return strings;
    }

    // Reads one <si>, concatenating the content of every <t> at any depth (plain, rich-text run,
    // phonetic run) in document order. The common case (a single <t>) never touches the builder.
    private static string ReadItem(XmlReader reader, StringBuilder builder)
    {
        if (reader.IsEmptyElement)
        {
            return string.Empty;
        }

        string? single = null;
        var multiple = false;

        reader.Read();

        while (!(reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "si"))
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t")
            {
                string text;

                if (reader.IsEmptyElement)
                {
                    text = string.Empty;
                    reader.Read(); // move past <t/>
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

        return multiple ? builder.ToString() : single ?? string.Empty;
    }
}
