using System.Globalization;
using System.Text;
using System.Text.Json;
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

// ---------------------------------------------------------------------------------------------------------
// K1FixtureBuilder — transforma o K1.xlsx (representado como JSON confidencial em samples/) num fixture
// ANONIMIZADO para o spike do grafo de dependências. Preserva a ESTRUTURA (refs, funções, nomes de sheet,
// o grafo de dependências) e troca os DADOS (conteúdo de string, números, valores de input, chaves de
// interpolação) por versões fake determinísticas. Produz:
//   samples/k1-anonymized.json  — referência humana (mesmo schema do original, já anonimizado)
//   samples/k1.myxl             — workbook MySheet salvo (o fixture de verdade)
// Uso: dotnet run -c Release --project tools/K1FixtureBuilder
// ---------------------------------------------------------------------------------------------------------

var root = FindRepoRoot();
var rawSheetPath = Path.Combine(root, "samples", "synthetic-sheet-as-json.json");
var rawInputPath = Path.Combine(root, "samples", "synthetic-input.json");
var outJsonPath = Path.Combine(root, "samples", "k1-anonymized.json");
var outMyxlPath = Path.Combine(root, "samples", "k1.myxl");

// Modo verificação: recarrega o .myxl e afirma a integridade estrutural do fixture.
if (args.Contains("--verify"))
{
    return Verify(outMyxlPath);
}

if (!File.Exists(rawSheetPath) || !File.Exists(rawInputPath))
{
    Console.Error.WriteLine($"Arquivos crus não encontrados em {Path.Combine(root, "samples")}.");
    return 1;
}

var sw = System.Diagnostics.Stopwatch.StartNew();
var anon = new Anonymizer();

// === 1. Input: mapa chave → valor (RawTimData) ===========================================================
Console.WriteLine("[1/6] Lendo input…");
var inputValues = LoadInput(rawInputPath);
Console.WriteLine($"      {inputValues.Count} chaves de input.");

// === 2. Sheet cru → coletar chaves de interpolação distintas ============================================
Console.WriteLine("[2/6] Parseando o sheet cru (97MB)…");
using var doc = JsonDocument.Parse(File.ReadAllText(rawSheetPath));
var rootEl = doc.RootElement;

Console.WriteLine("[3/6] Coletando chaves de interpolação…");
var keys = new SortedSet<string>(StringComparer.Ordinal);
foreach (var sheet in rootEl.GetProperty("sheets").EnumerateArray())
{
    foreach (var cell in sheet.GetProperty("cells").EnumerateArray())
    {
        if (cell.GetProperty("value").ValueKind == JsonValueKind.String)
        {
            CollectKeys(cell.GetProperty("value").GetString()!, keys);
        }
    }
}

// chave original → (linha 1-based no sheet Input, nome genérico IN0001…)
var keyRow = new Dictionary<string, int>(StringComparer.Ordinal);
var keyGeneric = new Dictionary<string, string>(StringComparer.Ordinal);
var row = 0;
foreach (var k in keys)
{
    row++;
    keyRow[k] = row;
    keyGeneric[k] = $"IN{row:0000}";
}
var keyCount = row;
Console.WriteLine($"      {keyCount} chaves de interpolação distintas.");

// Mapa de anonimização de nomes de sheet: nome de exibição → identificador seguro (Sheet_{n}).
// Seguro = simples (sem espaço, re-parseável sem aspas) e NÃO parecido com ref de célula (o '_' quebra o
// padrão letras+dígitos). Os qualificadores por codeName (Sheet8 etc.) não estão aqui e ficam intactos
// (já eram #REF! por não existir sheet com esse nome).
var sheetFake = new Dictionary<string, string>(StringComparer.Ordinal);
foreach (var sheet in rootEl.GetProperty("sheets").EnumerateArray())
{
    var sn = sheet.GetProperty("sheetName").GetString()!;
    sheetFake[sn] = $"Sheet_{sheet.GetProperty("sheetIndex").GetInt32() + 1}";
}

// === 3. Transformar todas as células (anonimizar + reescrever interpolações) ============================
Console.WriteLine("[4/6] Transformando células…");
var model = new List<SheetModel>();
var totalCells = 0;
var interpolationRewrites = 0;

foreach (var sheet in rootEl.GetProperty("sheets").EnumerateArray())
{
    var sm = new SheetModel(
        sheetFake[sheet.GetProperty("sheetName").GetString()!],
        sheet.GetProperty("sheetIndex").GetInt32()
    );

    foreach (var cell in sheet.GetProperty("cells").EnumerateArray())
    {
        var cref = cell.GetProperty("ref").GetString()!;
        var type = cell.GetProperty("type").GetString()!;
        var valueEl = cell.GetProperty("value");
        var value =
            valueEl.ValueKind == JsonValueKind.String ? valueEl.GetString()! : valueEl.ToString();

        CellModel transformed;

        if (value.Contains('{'))
        {
            // scalar/customExpression com interpolação → vira fórmula concatenando literais anon + Input!C{row}
            transformed = new CellModel(
                cref,
                CellKind.Formula,
                RewriteInterpolation(value, keyRow, anon)
            );
            interpolationRewrites++;
        }
        else if (type == "formula")
        {
            transformed = new CellModel(
                cref,
                CellKind.Formula,
                anon.RewriteFormula(value, sheetFake)
            );
        }
        else // scalar sem interpolação
        {
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var n))
            {
                transformed = new CellModel(cref, CellKind.Number, anon.Number(n));
            }
            else
            {
                transformed = new CellModel(cref, CellKind.Text, anon.Text(value));
            }
        }

        sm.Cells.Add(transformed);
        totalCells++;
    }

    model.Add(sm);
}

// === 4. Sheet "Input" (A: nome genérico, B: valor anon, C: XLOOKUP) =====================================
var inputSheet = new SheetModel("Input", model.Count);
foreach (var (k, r) in keyRow.OrderBy(p => p.Value))
{
    inputSheet.Cells.Add(new CellModel($"A{r}", CellKind.Text, keyGeneric[k]));

    // Coluna B: valor de input anonimizado, ou blank.
    if (inputValues.TryGetValue(k, out var iv) && !iv.IsBlank)
    {
        inputSheet.Cells.Add(
            iv.IsNumber
                ? new CellModel($"B{r}", CellKind.Number, anon.Number(iv.Num))
                : new CellModel($"B{r}", CellKind.Text, anon.Text(iv.Str!))
        );
    }

    // Coluna C: o valor resolvido da chave, referenciando B{r} DIRETAMENTE (não um XLOOKUP de coluna
    // inteira). A célula-alvo que as interpolações referenciam. Ref direta desacopla os inputs — editar
    // B{r} só afeta as interpolações da chave r (um XLOOKUP sobre $B$1:$B$N acoplaria todo input a tudo).
    inputSheet.Cells.Add(new CellModel($"C{r}", CellKind.Formula, $"=B{r}"));
}
model.Add(inputSheet);
Console.WriteLine(
    $"      {totalCells} células transformadas; {interpolationRewrites} interpolações reescritas; sheet Input com {inputSheet.Cells.Count} células."
);

// Libera o documento cru antes de construir/serializar.
doc.Dispose();

// === 5. Escrever o JSON anonimizado (referência humana) =================================================
Console.WriteLine("[5/6] Escrevendo JSON anonimizado…");
WriteAnonymizedJson(outJsonPath, model);

// === 6. Construir o Workbook e salvar .myxl =============================================================
Console.WriteLine("[6/6] Construindo Workbook e salvando .myxl…");
var (workbook, parseFailures) = BuildWorkbook(model, anon);
workbook.Save(outMyxlPath, new WorkbookSaveOptions { Compression = WorkbookCompression.Brotli });

sw.Stop();
Console.WriteLine();
Console.WriteLine("=== RESUMO ===");
Console.WriteLine($"Sheets:                 {model.Count} (26 originais + Input)");
Console.WriteLine($"Células (com Input):    {model.Sum(s => s.Cells.Count)}");
Console.WriteLine($"Interpolações→fórmula:  {interpolationRewrites}");
Console.WriteLine($"Chaves de input:        {keyCount}");
Console.WriteLine($"Falhas de parse:        {parseFailures} (substituídas por scalar aleatório)");
Console.WriteLine(
    $"JSON anonimizado:       {outJsonPath} ({new FileInfo(outJsonPath).Length / 1_000_000.0:F1} MB)"
);
Console.WriteLine(
    $"Fixture:                {outMyxlPath} ({new FileInfo(outMyxlPath).Length / 1_000_000.0:F1} MB)"
);
Console.WriteLine($"Tempo total:            {sw.Elapsed.TotalSeconds:F1}s");
return 0;

// ---------------------------------------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------------------------------------

static int Verify(string myxlPath)
{
    if (!File.Exists(myxlPath))
    {
        Console.Error.WriteLine($"Fixture não encontrado: {myxlPath}");
        return 1;
    }

    Console.WriteLine($"Carregando {myxlPath} …");
    var sw = System.Diagnostics.Stopwatch.StartNew();
    var wb = Workbook.Load(myxlPath);
    sw.Stop();
    Console.WriteLine($"Load: {sw.Elapsed.TotalSeconds:F1}s");

    var ok = true;
    void Check(string label, bool cond, string detail = "")
    {
        Console.WriteLine($"  [{(cond ? "OK" : "FALHA")}] {label} {detail}");
        ok &= cond;
    }

    Check("27 sheets", wb.Sheets.Count == 27, $"(={wb.Sheets.Count})");
    var hasInput = wb.TryGetSheet("Input", out var input);
    Check("sheet Input existe", hasInput);
    if (hasInput)
    {
        Check("Input tem 7731 células (2577×3)", input!.Count == 7731, $"(={input.Count})");
    }

    Workbook.RunWithLargeStack(() =>
    {
        // XLOOKUP resolve: para algum r, Input!C{r} == Input!B{r} (com B não-blank).
        var xlookupOk = false;
        for (var r = 1; r <= 10 && !xlookupOk; r++)
        {
            var c = wb.GetCellValue("Input", $"C{r}");
            var b = wb.GetCellValue("Input", $"B{r}");
            if (b.Kind != ComputedValueKind.Blank && Equals(c.AsObject(), b.AsObject()))
            {
                xlookupOk = true;
            }
        }
        Check("Input!Cr == Br (ref direta resolve)", xlookupOk);

        // Célula de interpolação conhecida: Sheet_1!S2 (ex-"Supplemental Statement", era {ENT_EIN}) → não é erro.
        var s2 = wb.GetCellValue("Sheet_1", "S2");
        Check(
            "S2 (interpolação) não é erro",
            s2.Kind != ComputedValueKind.Error,
            $"(kind={s2.Kind})"
        );

        // Nomes de sheet anonimizados: os originais não devem mais existir.
        Check(
            "nomes de sheet anonimizados",
            !wb.TryGetSheet("Supplemental Statement", out _)
                && !wb.TryGetSheet("Standard Footnotes 2025", out _)
        );

        // Amostra da sheet grande (Sheet_3 = ex-"Standard Footnotes 2025"): computa 2000 células sem crash.
        var big = wb["Sheet_3"];
        int errors = 0,
            evaluated = 0;
        foreach (var id in big.Keys.Take(2000))
        {
            var v = wb.GetCellValue("Sheet_3", id);
            evaluated++;
            if (v.Kind == ComputedValueKind.Error)
            {
                errors++;
            }
        }
        Console.WriteLine(
            $"  [info] {evaluated} células avaliadas na sheet grande; {errors} deram erro (#REF!/#DIV0 são normais)"
        );
        return 0;
    });

    Console.WriteLine(ok ? "\nVERIFICAÇÃO OK" : "\nVERIFICAÇÃO FALHOU");
    return ok ? 0 : 1;
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

static Dictionary<string, InputValue> LoadInput(string path)
{
    using var input = JsonDocument.Parse(File.ReadAllText(path));
    var map = new Dictionary<string, InputValue>(StringComparer.Ordinal);
    var data = input.RootElement.GetProperty("Data");
    foreach (var element in data.EnumerateArray())
    {
        if (!element.TryGetProperty("RawTimData", out var raw))
        {
            continue;
        }
        foreach (var prop in raw.EnumerateObject())
        {
            map[prop.Name] = prop.Value.ValueKind switch
            {
                JsonValueKind.Number => new InputValue(true, prop.Value.GetDouble(), null, false),
                JsonValueKind.String => new InputValue(
                    false,
                    0,
                    prop.Value.GetString(),
                    string.IsNullOrEmpty(prop.Value.GetString())
                ),
                JsonValueKind.Null => new InputValue(false, 0, null, true),
                _ => new InputValue(false, 0, prop.Value.ToString(), false),
            };
        }
    }
    return map;
}

static void CollectKeys(string value, SortedSet<string> keys)
{
    var i = 0;
    while (i < value.Length)
    {
        if (value[i] == '{')
        {
            var j = value.IndexOf('}', i);
            if (j < 0)
            {
                break;
            }
            keys.Add(value.Substring(i + 1, j - i - 1));
            i = j + 1;
        }
        else
        {
            i++;
        }
    }
}

// Reescreve uma célula com {Key} numa fórmula: literais viram "…" (anonimizados), chaves viram Input!C{row},
// tudo concatenado com &. Ex.: -{FDK1_USWTHTAX} → ="-"&Input!C123 ; {ENT_EIN} → =Input!C7
static string RewriteInterpolation(string value, Dictionary<string, int> keyRow, Anonymizer anon)
{
    var parts = new List<string>();
    var i = 0;
    while (i < value.Length)
    {
        if (value[i] == '{')
        {
            var j = value.IndexOf('}', i);
            if (j < 0)
            {
                // '{' sem par: trata o resto como literal
                var tail = value.Substring(i);
                if (tail.Length > 0)
                {
                    parts.Add(Quote(anon.Text(tail)));
                }
                break;
            }
            var key = value.Substring(i + 1, j - i - 1);
            parts.Add(
                keyRow.TryGetValue(key, out var r) ? $"Input!C{r}" : Quote(anon.Text($"{{{key}}}"))
            );
            i = j + 1;
        }
        else
        {
            var k = value.IndexOf('{', i);
            if (k < 0)
            {
                k = value.Length;
            }
            var literal = value.Substring(i, k - i);
            if (literal.Length > 0)
            {
                parts.Add(Quote(anon.Text(literal)));
            }
            i = k;
        }
    }

    return parts.Count == 0 ? "=\"\"" : "=" + string.Join("&", parts);
}

static string Quote(string s) => "\"" + s.Replace("\"", "\"\"") + "\"";

static void WriteAnonymizedJson(string path, List<SheetModel> model)
{
    using var stream = File.Create(path);
    using var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true });
    writer.WriteStartObject();
    writer.WriteString("fileName", "k1-anonymized.xlsx");
    writer.WritePropertyName("sheets");
    writer.WriteStartArray();
    foreach (var sheet in model)
    {
        writer.WriteStartObject();
        writer.WriteString("sheetName", sheet.Name);
        writer.WriteNumber("sheetIndex", sheet.Index);
        writer.WritePropertyName("cells");
        writer.WriteStartArray();
        foreach (var cell in sheet.Cells)
        {
            writer.WriteStartObject();
            writer.WriteString("ref", cell.Ref);
            writer.WriteString("type", cell.Kind == CellKind.Formula ? "formula" : "scalar");
            writer.WriteString("value", cell.Value);
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
        writer.WriteEndObject();
    }
    writer.WriteEndArray();
    writer.WriteEndObject();
}

static (Workbook, int) BuildWorkbook(List<SheetModel> model, Anonymizer anon)
{
    var workbook = new Workbook();
    var failures = 0;
    var rng = new Random(1234); // fallback determinístico para fórmulas que não parseiam

    foreach (var sm in model.OrderBy(s => s.Index))
    {
        var sheet = workbook.Sheets.Add(sm.Name);
        foreach (var cell in sm.Cells)
        {
            switch (cell.Kind)
            {
                case CellKind.Number:
                    sheet[cell.Ref] = new NumberValue(
                        double.Parse(cell.Value, CultureInfo.InvariantCulture)
                    );
                    break;
                case CellKind.Text:
                    sheet[cell.Ref] = new StringValue(cell.Value);
                    break;
                default: // Formula
                    try
                    {
                        sheet[cell.Ref] = ExpressionParser.Parse(cell.Value, sheet);
                    }
                    catch (ParseException)
                    {
                        // Fórmula estranha: substitui por um scalar aleatório (decisão do plano).
                        sheet[cell.Ref] = new NumberValue(rng.Next(0, 1000));
                        failures++;
                    }
                    break;
            }
        }
    }

    return (workbook, failures);
}

// ---------------------------------------------------------------------------------------------------------
// Tipos
// ---------------------------------------------------------------------------------------------------------

enum CellKind
{
    Number,
    Text,
    Formula,
}

sealed record CellModel(string Ref, CellKind Kind, string Value);

sealed class SheetModel(string name, int index)
{
    public string Name { get; } = name;
    public int Index { get; } = index;
    public List<CellModel> Cells { get; } = [];
}

readonly record struct InputValue(bool IsNumber, double Num, string? Str, bool IsBlank);

// Anonimização determinística: mesma entrada → mesma saída (preserva igualdade entre uma célula "Show" e o
// literal "Show" dentro de uma fórmula), mas irreconhecível. Preserva o SHAPE (letra→letra, dígito→dígito,
// tamanho), então refs/nomes de função/estrutura nunca são tocados por acidente.
sealed class Anonymizer
{
    private readonly Dictionary<string, string> _textCache = new(StringComparer.Ordinal);

    private static readonly string[] Lorem = (
        "lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore "
        + "et dolore magna aliqua enim ad minim veniam quis nostrud exercitation ullamco laboris nisi aliquip "
        + "ex ea commodo consequat duis aute irure in reprehenderit voluptate velit esse cillum fugiat nulla"
    ).Split(' ', StringSplitOptions.RemoveEmptyEntries);

    // Texto → lorem ipsum DETERMINÍSTICO (mesma origem → mesma saída, via cache): preserva igualdade de
    // graça (a célula "Show" e o literal "Show" numa fórmula viram o mesmo lorem, então IF/COUNTIF continuam
    // casando), com contagem de palavras ~ a do original. Claramente fake e irreconhecível.
    public string Text(string original)
    {
        if (original.Length == 0)
        {
            return original;
        }
        if (_textCache.TryGetValue(original, out var cached))
        {
            return cached;
        }

        var rng = new Random(unchecked((int)StableSeed(original)));
        var words = Math.Max(1, original.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length);
        var sb = new StringBuilder();
        for (var i = 0; i < words; i++)
        {
            if (i > 0)
            {
                sb.Append(' ');
            }
            sb.Append(Lorem[rng.Next(Lorem.Length)]);
        }

        var fake = sb.ToString();
        _textCache[original] = fake;
        return fake;
    }

    public string Number(double v)
    {
        var repr = v.ToString(CultureInfo.InvariantCulture);
        var rng = new Random(unchecked((int)StableSeed(repr)));
        var neg = v < 0;
        var abs = Math.Abs(v);
        var intDigits = abs < 1 ? 1 : (int)Math.Floor(Math.Log10(abs)) + 1;
        var mag = (long)Math.Pow(10, Math.Min(intDigits, 15));
        double fake = rng.NextInt64(0, Math.Max(mag, 1));
        if (v % 1 != 0)
        {
            fake += Math.Round(rng.NextDouble(), 2);
        }
        return (neg ? -fake : fake).ToString(CultureInfo.InvariantCulture);
    }

    // Reescreve uma fórmula anonimizando (1) o conteúdo de literais "…" (via o mapa determinístico) e
    // (2) os qualificadores de sheet (identificador ou 'nome' antes de '!') que casam com um nome de sheet
    // renomeado. Refs de célula, números e nomes de função ficam intactos. Qualificadores por codeName
    // (Sheet8 etc.) não estão no mapa e passam verbatim (já eram #REF!).
    public string RewriteFormula(string formula, IReadOnlyDictionary<string, string> sheetFake)
    {
        var sb = new StringBuilder(formula.Length);
        var i = 0;
        while (i < formula.Length)
        {
            var c = formula[i];
            if (c == '"')
            {
                sb.Append('"');
                i++;
                var lit = new StringBuilder();
                while (i < formula.Length)
                {
                    if (formula[i] == '"')
                    {
                        if (i + 1 < formula.Length && formula[i + 1] == '"')
                        {
                            lit.Append('"');
                            i += 2;
                            continue;
                        }
                        i++;
                        break;
                    }
                    lit.Append(formula[i]);
                    i++;
                }
                sb.Append(Text(lit.ToString()).Replace("\"", "\"\""));
                sb.Append('"');
            }
            else if (c == '\'')
            {
                // Nome de sheet entre aspas simples. Lê o conteúdo (com '' escapado); se seguido de '!' e
                // no mapa, emite o nome fake (simples → sem aspas), senão mantém verbatim.
                i++;
                var name = new StringBuilder();
                while (i < formula.Length)
                {
                    if (formula[i] == '\'')
                    {
                        if (i + 1 < formula.Length && formula[i + 1] == '\'')
                        {
                            name.Append('\'');
                            i += 2;
                            continue;
                        }
                        i++;
                        break;
                    }
                    name.Append(formula[i]);
                    i++;
                }
                var raw = name.ToString();
                if (
                    i < formula.Length
                    && formula[i] == '!'
                    && sheetFake.TryGetValue(raw, out var fakeQ)
                )
                {
                    sb.Append(fakeQ);
                }
                else
                {
                    sb.Append('\'').Append(raw.Replace("'", "''")).Append('\'');
                }
            }
            else if (char.IsLetter(c) || c is '_' or '$')
            {
                var start = i;
                while (
                    i < formula.Length
                    && (char.IsLetterOrDigit(formula[i]) || formula[i] is '_' or '.' or '$')
                )
                {
                    i++;
                }
                var token = formula[start..i];
                // Só reescreve quando é um qualificador de sheet renomeado (token seguido de '!').
                sb.Append(
                    i < formula.Length
                    && formula[i] == '!'
                    && sheetFake.TryGetValue(token, out var fake)
                        ? fake
                        : token
                );
            }
            else
            {
                sb.Append(c);
                i++;
            }
        }
        return sb.ToString();
    }

    // FNV-1a 32-bit: hash estável entre execuções (string.GetHashCode é randomizado por processo).
    private static uint StableSeed(string s)
    {
        var h = 2166136261u;
        foreach (var c in s)
        {
            h = (h ^ c) * 16777619u;
        }
        return h;
    }
}
