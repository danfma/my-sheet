using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using Danfma.MySheet;
using Danfma.MySheet.Excel;

// ---------------------------------------------------------------------------------------------------------
// SyntheticK1Builder — gera um fixture NÃO-confidencial com o perfil do K1 (~620k células, ~360k fórmulas)
// para os benchmarks de I/O Excel (--excel-memory), escrevendo o .xlsx DIRETAMENTE (ZipArchive + XmlWriter)
// para controlar shapes que o ExcelExport não produz: grupos de shared formula reais (<f t="shared" ref si>),
// shared strings com alta duplicação, rich text, xml:space="preserve", inline strings, booleans, erros e
// datas ISO (t="d"). Determinístico: mesma execução → mesmo arquivo. Produz:
//   samples/k1-synthetic.xlsx — a entrada dos cenários de LOAD do harness
//   samples/k1-synthetic.myxl — workbook MySheet salvo (fallback dos cenários export/merge do harness)
// Uso: dotnet run -c Release --project tools/SyntheticK1Builder
// ---------------------------------------------------------------------------------------------------------

const int DataRows = 100_000; // Data!A1:B100000 — 200k células numéricas
const int MainRows = 60_000; // Main linhas 2..60001
const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const string PkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

// Colunas de fórmula da sheet Main: 6 grupos de shared formula (si = índice), master na linha 2.
// Misturam refs cross-sheet, relativas, absolutas ($) e chamadas de função — o que o load real enfrenta.
(string Column, string MasterFormula)[] formulaGroups =
[
    ("B", "Data!A1*2+1"),
    ("C", "Data!B1+B2"),
    ("D", "IF(C2>0,C2*0.98,0)"),
    ("E", "ROUND(B2+C2,2)"),
    ("F", "MAX(B2,C2)"),
    ("G", "(B2+C2+D2)/3"),
];

var root = FindRepoRoot();
var samples = Path.Combine(root, "samples");
Directory.CreateDirectory(samples);
var xlsxPath = Path.Combine(samples, "k1-synthetic.xlsx");
var myxlPath = Path.Combine(samples, "k1-synthetic.myxl");

var sw = Stopwatch.StartNew();

Console.WriteLine("[1/3] Escrevendo k1-synthetic.xlsx (ZipArchive + XmlWriter)…");
var stats = WriteXlsx(xlsxPath, formulaGroups);
Console.WriteLine(
    $"      {stats.Cells:N0} células ({stats.Formulas:N0} fórmulas, {stats.TextCells:N0} texto), "
        + $"{stats.UniqueStrings} shared strings únicas — {new FileInfo(xlsxPath).Length / 1_000_000.0:F1} MB"
);

Console.WriteLine("[2/3] Recarregando via ExcelFile.Load (sanidade)…");
var workbook = ExcelFile.Load(xlsxPath);
var loadedCells = workbook.Sheets.Values.Sum(s => s.Count);
Console.WriteLine($"      {workbook.Sheets.Count} sheets, {loadedCells:N0} células carregadas.");

Console.WriteLine("[3/3] Salvando k1-synthetic.myxl (Brotli)…");
workbook.Save(myxlPath, new WorkbookSaveOptions { Compression = WorkbookCompression.Brotli });
Console.WriteLine($"      {new FileInfo(myxlPath).Length / 1_000_000.0:F1} MB");

sw.Stop();
Console.WriteLine();
Console.WriteLine("=== RESUMO ===");
Console.WriteLine($"xlsx:  {xlsxPath}");
Console.WriteLine($"myxl:  {myxlPath}");
Console.WriteLine($"Tempo: {sw.Elapsed.TotalSeconds:F1}s");
return 0;

// ---------------------------------------------------------------------------------------------------------
// Escrita do .xlsx
// ---------------------------------------------------------------------------------------------------------

static BuildStats WriteXlsx(string path, (string Column, string MasterFormula)[] formulaGroups)
{
    using var zip = new ZipArchive(File.Create(path), ZipArchiveMode.Create);
    var pool = BuildStringPool();
    var stats = new BuildStats { UniqueStrings = pool.Entries.Count };

    WriteEntry(zip, "[Content_Types].xml", WriteContentTypes);
    WriteEntry(zip, "_rels/.rels", WritePackageRels);
    WriteEntry(zip, "xl/workbook.xml", w => WriteWorkbook(w, DataRows));
    WriteEntry(zip, "xl/_rels/workbook.xml.rels", WriteWorkbookRels);
    WriteEntry(zip, "xl/styles.xml", WriteStyles);
    WriteEntry(zip, "xl/worksheets/sheet1.xml", w => WriteDataSheet(w, stats));
    WriteEntry(zip, "xl/worksheets/sheet2.xml", w => WriteMainSheet(w, pool, formulaGroups, stats));
    WriteEntry(zip, "xl/sharedStrings.xml", w => WriteSharedStrings(w, pool, stats));

    return stats;
}

static void WriteEntry(ZipArchive zip, string name, Action<XmlWriter> write)
{
    var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
    using var stream = entry.Open();
    using var writer = XmlWriter.Create(
        stream,
        new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false }
    );
    write(writer);
}

static void WriteContentTypes(XmlWriter w)
{
    const string ns = "http://schemas.openxmlformats.org/package/2006/content-types";
    w.WriteStartElement("Types", ns);

    w.WriteStartElement("Default", ns);
    w.WriteAttributeString("Extension", "rels");
    w.WriteAttributeString(
        "ContentType",
        "application/vnd.openxmlformats-package.relationships+xml"
    );
    w.WriteEndElement();

    w.WriteStartElement("Default", ns);
    w.WriteAttributeString("Extension", "xml");
    w.WriteAttributeString("ContentType", "application/xml");
    w.WriteEndElement();

    WriteOverride(
        w,
        ns,
        "/xl/workbook.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    );
    WriteOverride(
        w,
        ns,
        "/xl/worksheets/sheet1.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    );
    WriteOverride(
        w,
        ns,
        "/xl/worksheets/sheet2.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    );
    WriteOverride(
        w,
        ns,
        "/xl/sharedStrings.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
    );
    WriteOverride(
        w,
        ns,
        "/xl/styles.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
    );

    w.WriteEndElement();

    static void WriteOverride(XmlWriter w, string ns, string part, string contentType)
    {
        w.WriteStartElement("Override", ns);
        w.WriteAttributeString("PartName", part);
        w.WriteAttributeString("ContentType", contentType);
        w.WriteEndElement();
    }
}

static void WritePackageRels(XmlWriter w)
{
    w.WriteStartElement("Relationships", PkgRelNs);
    w.WriteStartElement("Relationship", PkgRelNs);
    w.WriteAttributeString("Id", "rId1");
    w.WriteAttributeString("Type", RelNs + "/officeDocument");
    w.WriteAttributeString("Target", "xl/workbook.xml");
    w.WriteEndElement();
    w.WriteEndElement();
}

static void WriteWorkbookRels(XmlWriter w)
{
    w.WriteStartElement("Relationships", PkgRelNs);
    WriteRel(w, "rId1", "/worksheet", "worksheets/sheet1.xml");
    WriteRel(w, "rId2", "/worksheet", "worksheets/sheet2.xml");
    WriteRel(w, "rId3", "/sharedStrings", "sharedStrings.xml");
    WriteRel(w, "rId4", "/styles", "styles.xml");
    w.WriteEndElement();

    static void WriteRel(XmlWriter w, string id, string typeSuffix, string target)
    {
        w.WriteStartElement("Relationship", PkgRelNs);
        w.WriteAttributeString("Id", id);
        w.WriteAttributeString("Type", RelNs + typeSuffix);
        w.WriteAttributeString("Target", target);
        w.WriteEndElement();
    }
}

static void WriteWorkbook(XmlWriter w, int dataRows)
{
    w.WriteStartElement("workbook", MainNs);
    w.WriteAttributeString("xmlns", "r", null, RelNs);

    w.WriteStartElement("sheets", MainNs);
    WriteSheet(w, "Data", 1, "rId1");
    WriteSheet(w, "Main", 2, "rId2");
    w.WriteEndElement();

    // Defined names exercitam LoadDefinedNames: uma constante e um range absoluto cross-sheet.
    w.WriteStartElement("definedNames", MainNs);
    WriteDefinedName(w, "TaxRate", "0.21");
    WriteDefinedName(w, "DataColA", $"Data!$A$1:$A${dataRows}");
    w.WriteEndElement();

    w.WriteEndElement();

    static void WriteSheet(XmlWriter w, string name, int sheetId, string relId)
    {
        w.WriteStartElement("sheet", MainNs);
        w.WriteAttributeString("name", name);
        w.WriteAttributeString("sheetId", sheetId.ToString(CultureInfo.InvariantCulture));
        w.WriteAttributeString("id", RelNs, relId);
        w.WriteEndElement();
    }

    static void WriteDefinedName(XmlWriter w, string name, string refersTo)
    {
        w.WriteStartElement("definedName", MainNs);
        w.WriteAttributeString("name", name);
        w.WriteString(refersTo);
        w.WriteEndElement();
    }
}

static void WriteStyles(XmlWriter w)
{
    // O mínimo aceito pelos leitores: uma entrada default em cada coleção obrigatória.
    w.WriteStartElement("styleSheet", MainNs);
    WriteSingleton(w, "fonts", "font");
    WriteSingleton(w, "fills", "fill");
    WriteSingleton(w, "borders", "border");
    WriteSingleton(w, "cellStyleXfs", "xf");
    WriteSingleton(w, "cellXfs", "xf");
    w.WriteEndElement();

    static void WriteSingleton(XmlWriter w, string collection, string item)
    {
        w.WriteStartElement(collection, MainNs);
        w.WriteAttributeString("count", "1");
        w.WriteStartElement(item, MainNs);
        w.WriteEndElement();
        w.WriteEndElement();
    }
}

// Data: duas colunas numéricas densas (A = inteiros pequenos, muito comuns em planilhas de negócio;
// B = decimais), determinísticas.
static void WriteDataSheet(XmlWriter w, BuildStats stats)
{
    var rng = new Random(20260710);

    w.WriteStartElement("worksheet", MainNs);
    w.WriteStartElement("dimension", MainNs);
    w.WriteAttributeString("ref", $"A1:B{DataRows}");
    w.WriteEndElement();
    w.WriteStartElement("sheetData", MainNs);

    for (var r = 1; r <= DataRows; r++)
    {
        w.WriteStartElement("row", MainNs);
        w.WriteAttributeString("r", r.ToString(CultureInfo.InvariantCulture));

        WriteNumberCell(w, $"A{r}", rng.Next(0, 1000).ToString(CultureInfo.InvariantCulture));
        WriteNumberCell(
            w,
            $"B{r}",
            Math.Round(rng.NextDouble() * 10_000, 2).ToString(CultureInfo.InvariantCulture)
        );
        stats.Cells += 2;

        w.WriteEndElement();
    }

    w.WriteEndElement();
    w.WriteEndElement();
}

// Main: coluna A de texto com alta duplicação (shared strings), colunas B..G em grupos de shared formula
// (master na linha 2 com ref+si, escravas só com si) e um bloco final de edge cases.
static void WriteMainSheet(
    XmlWriter w,
    StringPool pool,
    (string Column, string MasterFormula)[] formulaGroups,
    BuildStats stats
)
{
    var lastFormulaRow = MainRows + 1; // linhas 2..60001
    var edgeRow = lastFormulaRow + 1;

    w.WriteStartElement("worksheet", MainNs);
    w.WriteStartElement("dimension", MainNs);
    w.WriteAttributeString("ref", $"A1:G{edgeRow}");
    w.WriteEndElement();
    w.WriteStartElement("sheetData", MainNs);

    // Linha 1: headers (texto).
    w.WriteStartElement("row", MainNs);
    w.WriteAttributeString("r", "1");
    for (var c = 0; c <= formulaGroups.Length; c++)
    {
        WriteSharedStringCell(w, $"{(char)('A' + c)}1", pool.HeaderIndex + c, stats);
        stats.Cells++;
    }
    w.WriteEndElement();

    for (var r = 2; r <= lastFormulaRow; r++)
    {
        w.WriteStartElement("row", MainNs);
        w.WriteAttributeString("r", r.ToString(CultureInfo.InvariantCulture));

        // A: label com duplicação alta; a cada 10k linhas, a entrada rich-text.
        var poolIndex = r % 10_000 == 0 ? pool.RichTextIndex : (r * 31) % pool.LabelCount;
        WriteSharedStringCell(w, $"A{r}", poolIndex, stats);
        stats.Cells++;
        stats.TextCells++;

        // B..G: master (linha 2) carrega texto + ref do grupo; escravas só o si.
        for (var g = 0; g < formulaGroups.Length; g++)
        {
            var (column, masterFormula) = formulaGroups[g];
            w.WriteStartElement("c", MainNs);
            w.WriteAttributeString("r", $"{column}{r}");
            w.WriteStartElement("f", MainNs);
            w.WriteAttributeString("t", "shared");
            w.WriteAttributeString("si", g.ToString(CultureInfo.InvariantCulture));
            if (r == 2)
            {
                w.WriteAttributeString("ref", $"{column}2:{column}{lastFormulaRow}");
                w.WriteString(masterFormula);
            }
            w.WriteEndElement();
            // Cached value (ignorado pelo MySheet, usado por leitores que não recalculam).
            w.WriteStartElement("v", MainNs);
            w.WriteString("0");
            w.WriteEndElement();
            w.WriteEndElement();
            stats.Cells++;
            stats.Formulas++;
        }

        w.WriteEndElement();
    }

    // Bloco final de edge cases: inline string com espaços, boolean, erro, data ISO e fórmula não-shared.
    w.WriteStartElement("row", MainNs);
    w.WriteAttributeString("r", edgeRow.ToString(CultureInfo.InvariantCulture));

    w.WriteStartElement("c", MainNs);
    w.WriteAttributeString("r", $"A{edgeRow}");
    w.WriteAttributeString("t", "inlineStr");
    w.WriteStartElement("is", MainNs);
    w.WriteStartElement("t", MainNs);
    w.WriteAttributeString("xml", "space", null, "preserve");
    w.WriteString(" inline  text ");
    w.WriteEndElement();
    w.WriteEndElement();
    w.WriteEndElement();

    w.WriteStartElement("c", MainNs);
    w.WriteAttributeString("r", $"B{edgeRow}");
    w.WriteAttributeString("t", "b");
    w.WriteStartElement("v", MainNs);
    w.WriteString("1");
    w.WriteEndElement();
    w.WriteEndElement();

    w.WriteStartElement("c", MainNs);
    w.WriteAttributeString("r", $"C{edgeRow}");
    w.WriteAttributeString("t", "e");
    w.WriteStartElement("v", MainNs);
    w.WriteString("#DIV/0!");
    w.WriteEndElement();
    w.WriteEndElement();

    w.WriteStartElement("c", MainNs);
    w.WriteAttributeString("r", $"D{edgeRow}");
    w.WriteAttributeString("t", "d");
    w.WriteStartElement("v", MainNs);
    w.WriteString("2026-07-10");
    w.WriteEndElement();
    w.WriteEndElement();

    w.WriteStartElement("c", MainNs);
    w.WriteAttributeString("r", $"E{edgeRow}");
    w.WriteStartElement("f", MainNs);
    w.WriteString("SUM(B2:B10)");
    w.WriteEndElement();
    w.WriteStartElement("v", MainNs);
    w.WriteString("0");
    w.WriteEndElement();
    w.WriteEndElement();

    stats.Cells += 5;
    stats.Formulas++;

    w.WriteEndElement(); // row
    w.WriteEndElement(); // sheetData
    w.WriteEndElement(); // worksheet
}

static void WriteNumberCell(XmlWriter w, string reference, string value)
{
    w.WriteStartElement("c", MainNs);
    w.WriteAttributeString("r", reference);
    w.WriteStartElement("v", MainNs);
    w.WriteString(value);
    w.WriteEndElement();
    w.WriteEndElement();
}

static void WriteSharedStringCell(XmlWriter w, string reference, int index, BuildStats stats)
{
    w.WriteStartElement("c", MainNs);
    w.WriteAttributeString("r", reference);
    w.WriteAttributeString("t", "s");
    w.WriteStartElement("v", MainNs);
    w.WriteString(index.ToString(CultureInfo.InvariantCulture));
    w.WriteEndElement();
    w.WriteEndElement();
    stats.StringRefs++;
}

static void WriteSharedStrings(XmlWriter w, StringPool pool, BuildStats stats)
{
    w.WriteStartElement("sst", MainNs);
    w.WriteAttributeString("count", stats.StringRefs.ToString(CultureInfo.InvariantCulture));
    w.WriteAttributeString(
        "uniqueCount",
        pool.Entries.Count.ToString(CultureInfo.InvariantCulture)
    );

    for (var i = 0; i < pool.Entries.Count; i++)
    {
        w.WriteStartElement("si", MainNs);

        if (i == pool.RichTextIndex)
        {
            // Rich text: dois runs; o segundo com espaços preservados (paridade com InnerText do DOM).
            w.WriteStartElement("r", MainNs);
            w.WriteStartElement("rPr", MainNs);
            w.WriteStartElement("b", MainNs);
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("t", MainNs);
            w.WriteString("synthetic");
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("r", MainNs);
            w.WriteStartElement("t", MainNs);
            w.WriteAttributeString("xml", "space", null, "preserve");
            w.WriteString(" rich text ");
            w.WriteEndElement();
            w.WriteEndElement();
        }
        else
        {
            var text = pool.Entries[i];
            w.WriteStartElement("t", MainNs);
            if (text.Length > 0 && (text[0] == ' ' || text[^1] == ' '))
            {
                w.WriteAttributeString("xml", "space", null, "preserve");
            }
            w.WriteString(text);
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    w.WriteEndElement();
}

// Pool de strings: 200 labels (195 simples + 5 com espaços nas pontas), 1 rich text e 7 headers.
static StringPool BuildStringPool()
{
    var entries = new List<string>();
    for (var i = 0; i < 195; i++)
    {
        entries.Add($"lorem label {i:000}");
    }
    for (var i = 0; i < 5; i++)
    {
        entries.Add($"  padded label {i:000}  ");
    }

    var richTextIndex = entries.Count;
    entries.Add("synthetic rich text "); // placeholder: o writer emite os runs; o texto é o InnerText esperado

    var headerIndex = entries.Count;
    foreach (var c in "ABCDEFG")
    {
        entries.Add($"Header {c}");
    }

    return new StringPool
    {
        Entries = entries,
        LabelCount = 200,
        RichTextIndex = richTextIndex,
        HeaderIndex = headerIndex,
    };
}

static string FindRepoRoot()
{
    var dir = AppContext.BaseDirectory;
    while (dir is not null && !File.Exists(Path.Combine(dir, "Danfma.MySheet.slnx")))
    {
        dir = Directory.GetParent(dir)?.FullName;
    }
    return dir ?? Directory.GetCurrentDirectory();
}

sealed class BuildStats
{
    public long Cells;
    public long Formulas;
    public long TextCells;
    public long StringRefs;
    public int UniqueStrings;
}

sealed class StringPool
{
    public required List<string> Entries { get; init; }
    public required int LabelCount { get; init; }
    public required int RichTextIndex { get; init; }
    public required int HeaderIndex { get; init; }
}
