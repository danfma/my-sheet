using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using Aspose.Cells;
using AsposeWorkbook = Aspose.Cells.Workbook;
using AsposeCell = Aspose.Cells.Cell;

namespace Danfma.MySheet.Benchmark;

// A/B allocation-fishing suite (plans/function-allocation-fishing.md, Phase 1). For each formula FAMILY a pair
// {Family}_MySheet × {Family}_Aspose (Aspose is Baseline=true) measures the cost of RE-EVALUATING one target
// formula from scratch, on synthetic data shaped after the real load. The MemoryDiagnoser is the allocation
// judge; the family with the worst MySheet Allocated (weighted by how often it shows up in the real load)
// becomes the first fishing target in Phase 2. NO production code is touched here — Phase 1 only measures.
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --filter *FunctionBenchmarks* --job short
//
// HOW REPEATED EVALUATION IS MEASURED (and why it is honest on both sides)
// -----------------------------------------------------------------------
// Each pair owns its OWN workbook holding data cells (plain values) plus exactly ONE target-formula cell. The
// goal is to measure the EVALUATION of that ONE target formula over its inputs — NOT the cost of rebuilding the
// data store. So the inputs are warmed once and stay resident; only the target is recomputed each invocation.
//
//   MySheet: the target Expression is evaluated DIRECTLY — targetExpr.Evaluate(new EvaluationContext(wb, sheet,
//            cell)) — exactly as Workbook.EvaluateCell does internally, but WITHOUT memoizing the target result.
//            The input value cache is warmed once in GlobalSetup (a single GetCellValue of the target), so every
//            range/cell read the target makes is a warm HIT (no SetDense, no dense-page allocation). What remains
//            allocated per call is precisely the target's own evaluation transient (criteria lists, pairwise
//            range buffers, materialized array vectors, new strings) — the thing Phase 2 fishes.
//            NOTE: InvalidateCache()+GetCellValue was the plan's EXAMPLE, but InvalidateCache drops the dense
//            store, forcing a 50k-cell data-store REBUILD every iteration — plumbing that swamps the formula
//            transient and is unfair to Aspose (whose data cells persist across a recompute). Direct evaluation
//            over a warm input cache is the honest, intent-aligned form ("recompute only the target"); the
//            Scalar-control anchor collapsing to ~0 alloc confirms the harness isolates the evaluation.
//
//   Aspose:  CalculateFormula() with ForceFullCalculation=true and EnableCalculationChain=false. As documented
//            in AsposeCompareHarness, the default cached calc-chain serves a repeated CalculateFormula() in
//            ~0 ms recomputing NOTHING; forcing full calculation with the chain disabled makes every call
//            recompute the whole (single-formula) set from scratch, reading its already-materialized (persistent,
//            warm) data cells. Because each workbook holds ONLY the target formula, "recompute everything" ==
//            "recompute the target" over warm inputs — the honest analogue of MySheet's direct re-evaluation.
//            Neither side rebuilds its data store; neither side is allowed to serve a stale cache.
//
// Nothing is ever saved on either side (the K1 lesson: a fair in-memory comparison keeps both engines fully in
// memory; no save means Aspose never injects its evaluation watermark and has no row/cell cap). A license is
// loaded only if ASPOSE_LICENSE_PATH points at an existing file; absent it, evaluation mode is sufficient for
// this in-memory scenario. GlobalSetup asserts, per pair, that BOTH engines produce the SAME value — a
// benchmark comparing engines that compute different things would be a lie.
//
// Data sizes: SmallN (~200 rows) and LargeN (~50k rows) for the range-consuming families; scalar/single-cell
// families (IF-passthrough, OR, text, scalar control) use one shape (size is irrelevant to them).
[MemoryDiagnoser]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
[CategoriesColumn]
public class FunctionBenchmarks
{
    private const int SmallN = 200;
    private const int LargeN = 50_000;

    // Each pair: a MySheet target (workbook + sheet + cell + parsed expression) and its Aspose mirror.
    private sealed class MyTarget
    {
        public required Workbook Workbook;
        public required string Sheet;
        public required string Cell;
        public required Expression Expression;

        // Direct evaluation of the target — the input value cache is already warm, so this recomputes ONLY the
        // target's transient without a data-store rebuild and without memoizing the result.
        public object? Eval() =>
            Expression.Evaluate(new EvaluationContext(Workbook, Sheet, Cell)).AsObject();

        // Populate the input value cache once so every subsequent Eval() reads its inputs as warm cache hits.
        public void Warm() => Workbook.GetCellValue(Sheet, Cell);
    }

    private sealed class AsTarget
    {
        public required AsposeWorkbook Workbook;
        public required AsposeCell Cell;

        public object? Eval()
        {
            Workbook.CalculateFormula();
            return Cell.Value;
        }
    }

    // --- One field per family/size pair -------------------------------------------------------------------
    private MyTarget _sumifsSmallMy = null!, _sumifsLargeMy = null!;
    private AsTarget _sumifsSmallAs = null!, _sumifsLargeAs = null!;

    private MyTarget _sumproductSmallMy = null!, _sumproductLargeMy = null!;
    private AsTarget _sumproductSmallAs = null!, _sumproductLargeAs = null!;

    private MyTarget _countifOpenSmallMy = null!, _countifOpenLargeMy = null!;
    private AsTarget _countifOpenSmallAs = null!, _countifOpenLargeAs = null!;

    private MyTarget _countifClosedSmallMy = null!, _countifClosedLargeMy = null!;
    private AsTarget _countifClosedSmallAs = null!, _countifClosedLargeAs = null!;

    private MyTarget _arraySmallMy = null!, _arrayLargeMy = null!;
    private AsTarget _arraySmallAs = null!, _arrayLargeAs = null!;

    private MyTarget _xlookupSmallMy = null!, _xlookupLargeMy = null!;
    private AsTarget _xlookupSmallAs = null!, _xlookupLargeAs = null!;

    private MyTarget _matchIntSmallMy = null!, _matchIntLargeMy = null!;
    private AsTarget _matchIntSmallAs = null!, _matchIntLargeAs = null!;

    private MyTarget _ifPassthroughMy = null!;
    private AsTarget _ifPassthroughAs = null!;

    private MyTarget _orDisjunctsMy = null!;
    private AsTarget _orDisjunctsAs = null!;

    private MyTarget _textUpperTrimMy = null!;
    private AsTarget _textUpperTrimAs = null!;

    private MyTarget _textLeftFindMy = null!;
    private AsTarget _textLeftFindAs = null!;

    private MyTarget _textConcatMy = null!;
    private AsTarget _textConcatAs = null!;

    private MyTarget _scalarControlMy = null!;
    private AsTarget _scalarControlAs = null!;

    [GlobalSetup]
    public void Setup()
    {
        TryLoadLicense();

        (_sumifsSmallMy, _sumifsSmallAs) = BuildSumifs(SmallN, "SUMIFS/200");
        (_sumifsLargeMy, _sumifsLargeAs) = BuildSumifs(LargeN, "SUMIFS/50k");

        (_sumproductSmallMy, _sumproductSmallAs) = BuildSumproduct(SmallN, "SUMPRODUCT/200");
        (_sumproductLargeMy, _sumproductLargeAs) = BuildSumproduct(LargeN, "SUMPRODUCT/50k");

        (_countifOpenSmallMy, _countifOpenSmallAs) = BuildCountif(SmallN, open: true, "COUNTIF-open/200");
        (_countifOpenLargeMy, _countifOpenLargeAs) = BuildCountif(LargeN, open: true, "COUNTIF-open/50k");

        (_countifClosedSmallMy, _countifClosedSmallAs) = BuildCountif(SmallN, open: false, "COUNTIF-closed/200");
        (_countifClosedLargeMy, _countifClosedLargeAs) = BuildCountif(LargeN, open: false, "COUNTIF-closed/50k");

        (_arraySmallMy, _arraySmallAs) = BuildArray(SmallN, "Array-INDEX-SMALL/200");
        (_arrayLargeMy, _arrayLargeAs) = BuildArray(LargeN, "Array-INDEX-SMALL/50k");

        (_xlookupSmallMy, _xlookupSmallAs) = BuildXlookup(SmallN, "XLOOKUP/200");
        (_xlookupLargeMy, _xlookupLargeAs) = BuildXlookup(LargeN, "XLOOKUP/50k");

        (_matchIntSmallMy, _matchIntSmallAs) = BuildMatchInt(SmallN, "MATCH-INT/200");
        (_matchIntLargeMy, _matchIntLargeAs) = BuildMatchInt(LargeN, "MATCH-INT/50k");

        (_ifPassthroughMy, _ifPassthroughAs) = BuildIfPassthrough("IF-passthrough");
        (_orDisjunctsMy, _orDisjunctsAs) = BuildOrDisjuncts("OR-disjuncts");
        (_textUpperTrimMy, _textUpperTrimAs) = BuildTextUpperTrim("Text-UPPER-TRIM");
        (_textLeftFindMy, _textLeftFindAs) = BuildTextLeftFind("Text-LEFT-FIND");
        (_textConcatMy, _textConcatAs) = BuildTextConcat("Text-concat");
        (_scalarControlMy, _scalarControlAs) = BuildScalarControl("Scalar-control");
    }

    // ==================================================================================================
    // Benchmarks — one pair per category. Aspose is the baseline; the ratio columns hang off it.
    // ==================================================================================================

    [BenchmarkCategory("SUMIFS/200"), Benchmark]
    public object? Sumifs200_MySheet() => _sumifsSmallMy.Eval();

    [BenchmarkCategory("SUMIFS/200"), Benchmark(Baseline = true)]
    public object? Sumifs200_Aspose() => _sumifsSmallAs.Eval();

    [BenchmarkCategory("SUMIFS/50k"), Benchmark]
    public object? Sumifs50k_MySheet() => _sumifsLargeMy.Eval();

    [BenchmarkCategory("SUMIFS/50k"), Benchmark(Baseline = true)]
    public object? Sumifs50k_Aspose() => _sumifsLargeAs.Eval();

    [BenchmarkCategory("SUMPRODUCT/200"), Benchmark]
    public object? Sumproduct200_MySheet() => _sumproductSmallMy.Eval();

    [BenchmarkCategory("SUMPRODUCT/200"), Benchmark(Baseline = true)]
    public object? Sumproduct200_Aspose() => _sumproductSmallAs.Eval();

    [BenchmarkCategory("SUMPRODUCT/50k"), Benchmark]
    public object? Sumproduct50k_MySheet() => _sumproductLargeMy.Eval();

    [BenchmarkCategory("SUMPRODUCT/50k"), Benchmark(Baseline = true)]
    public object? Sumproduct50k_Aspose() => _sumproductLargeAs.Eval();

    [BenchmarkCategory("COUNTIF-open/200"), Benchmark]
    public object? CountifOpen200_MySheet() => _countifOpenSmallMy.Eval();

    [BenchmarkCategory("COUNTIF-open/200"), Benchmark(Baseline = true)]
    public object? CountifOpen200_Aspose() => _countifOpenSmallAs.Eval();

    [BenchmarkCategory("COUNTIF-open/50k"), Benchmark]
    public object? CountifOpen50k_MySheet() => _countifOpenLargeMy.Eval();

    [BenchmarkCategory("COUNTIF-open/50k"), Benchmark(Baseline = true)]
    public object? CountifOpen50k_Aspose() => _countifOpenLargeAs.Eval();

    [BenchmarkCategory("COUNTIF-closed/200"), Benchmark]
    public object? CountifClosed200_MySheet() => _countifClosedSmallMy.Eval();

    [BenchmarkCategory("COUNTIF-closed/200"), Benchmark(Baseline = true)]
    public object? CountifClosed200_Aspose() => _countifClosedSmallAs.Eval();

    [BenchmarkCategory("COUNTIF-closed/50k"), Benchmark]
    public object? CountifClosed50k_MySheet() => _countifClosedLargeMy.Eval();

    [BenchmarkCategory("COUNTIF-closed/50k"), Benchmark(Baseline = true)]
    public object? CountifClosed50k_Aspose() => _countifClosedLargeAs.Eval();

    [BenchmarkCategory("Array-INDEX-SMALL/200"), Benchmark]
    public object? Array200_MySheet() => _arraySmallMy.Eval();

    [BenchmarkCategory("Array-INDEX-SMALL/200"), Benchmark(Baseline = true)]
    public object? Array200_Aspose() => _arraySmallAs.Eval();

    [BenchmarkCategory("Array-INDEX-SMALL/50k"), Benchmark]
    public object? Array50k_MySheet() => _arrayLargeMy.Eval();

    [BenchmarkCategory("Array-INDEX-SMALL/50k"), Benchmark(Baseline = true)]
    public object? Array50k_Aspose() => _arrayLargeAs.Eval();

    [BenchmarkCategory("XLOOKUP/200"), Benchmark]
    public object? Xlookup200_MySheet() => _xlookupSmallMy.Eval();

    [BenchmarkCategory("XLOOKUP/200"), Benchmark(Baseline = true)]
    public object? Xlookup200_Aspose() => _xlookupSmallAs.Eval();

    [BenchmarkCategory("XLOOKUP/50k"), Benchmark]
    public object? Xlookup50k_MySheet() => _xlookupLargeMy.Eval();

    [BenchmarkCategory("XLOOKUP/50k"), Benchmark(Baseline = true)]
    public object? Xlookup50k_Aspose() => _xlookupLargeAs.Eval();

    [BenchmarkCategory("MATCH-INT/200"), Benchmark]
    public object? MatchInt200_MySheet() => _matchIntSmallMy.Eval();

    [BenchmarkCategory("MATCH-INT/200"), Benchmark(Baseline = true)]
    public object? MatchInt200_Aspose() => _matchIntSmallAs.Eval();

    [BenchmarkCategory("MATCH-INT/50k"), Benchmark]
    public object? MatchInt50k_MySheet() => _matchIntLargeMy.Eval();

    [BenchmarkCategory("MATCH-INT/50k"), Benchmark(Baseline = true)]
    public object? MatchInt50k_Aspose() => _matchIntLargeAs.Eval();

    [BenchmarkCategory("IF-passthrough"), Benchmark]
    public object? IfPassthrough_MySheet() => _ifPassthroughMy.Eval();

    [BenchmarkCategory("IF-passthrough"), Benchmark(Baseline = true)]
    public object? IfPassthrough_Aspose() => _ifPassthroughAs.Eval();

    [BenchmarkCategory("OR-disjuncts"), Benchmark]
    public object? OrDisjuncts_MySheet() => _orDisjunctsMy.Eval();

    [BenchmarkCategory("OR-disjuncts"), Benchmark(Baseline = true)]
    public object? OrDisjuncts_Aspose() => _orDisjunctsAs.Eval();

    [BenchmarkCategory("Text-UPPER-TRIM"), Benchmark]
    public object? TextUpperTrim_MySheet() => _textUpperTrimMy.Eval();

    [BenchmarkCategory("Text-UPPER-TRIM"), Benchmark(Baseline = true)]
    public object? TextUpperTrim_Aspose() => _textUpperTrimAs.Eval();

    [BenchmarkCategory("Text-LEFT-FIND"), Benchmark]
    public object? TextLeftFind_MySheet() => _textLeftFindMy.Eval();

    [BenchmarkCategory("Text-LEFT-FIND"), Benchmark(Baseline = true)]
    public object? TextLeftFind_Aspose() => _textLeftFindAs.Eval();

    [BenchmarkCategory("Text-concat"), Benchmark]
    public object? TextConcat_MySheet() => _textConcatMy.Eval();

    [BenchmarkCategory("Text-concat"), Benchmark(Baseline = true)]
    public object? TextConcat_Aspose() => _textConcatAs.Eval();

    [BenchmarkCategory("Scalar-control"), Benchmark]
    public object? ScalarControl_MySheet() => _scalarControlMy.Eval();

    [BenchmarkCategory("Scalar-control"), Benchmark(Baseline = true)]
    public object? ScalarControl_Aspose() => _scalarControlAs.Eval();

    // ==================================================================================================
    // Builders — each returns a (MySheet, Aspose) pair populated from the SAME deterministic data, wiring the
    // target formula and asserting cross-engine value equivalence before returning.
    // ==================================================================================================

    private const string SheetName = "S";
    private const string Target = "Z1"; // a cell no open-range family references

    // SUMIFS with 3 criteria pairs: SUM(amount) where group="G3" AND region="R2" AND amount>=50.
    // A=group text, B=region text, C=amount number.
    private static (MyTarget, AsTarget) BuildSumifs(int n, string name)
    {
        var (my, sheet) = NewMy();
        var (asWb, asWs) = NewAspose();
        var last = n + 1;

        for (var r = 2; r <= last; r++)
        {
            var group = "G" + (r % 10);
            var region = "R" + (r % 5);
            var amount = (double)((r % 100) + 1);

            sheet[$"A{r}"] = new StringValue(group);
            sheet[$"B{r}"] = new StringValue(region);
            sheet[$"C{r}"] = new NumberValue(amount);

            asWs.Cells[$"A{r}"].PutValue(group);
            asWs.Cells[$"B{r}"].PutValue(region);
            asWs.Cells[$"C{r}"].PutValue(amount);
        }

        var formula = $"=SUMIFS(C2:C{last},A2:A{last},\"G3\",B2:B{last},\"R2\",C2:C{last},\">=50\")";
        return Finish(my, sheet, asWb, asWs, formula, name);
    }

    // SUMPRODUCT of three numeric columns.
    private static (MyTarget, AsTarget) BuildSumproduct(int n, string name)
    {
        var (my, sheet) = NewMy();
        var (asWb, asWs) = NewAspose();
        var last = n + 1;

        for (var r = 2; r <= last; r++)
        {
            double a = (r % 7) + 1, b = (r % 5) + 1, c = (r % 3) + 1;
            sheet[$"A{r}"] = new NumberValue(a);
            sheet[$"B{r}"] = new NumberValue(b);
            sheet[$"C{r}"] = new NumberValue(c);
            asWs.Cells[$"A{r}"].PutValue(a);
            asWs.Cells[$"B{r}"].PutValue(b);
            asWs.Cells[$"C{r}"].PutValue(c);
        }

        var formula = $"=SUMPRODUCT(A2:A{last},B2:B{last},C2:C{last})";
        return Finish(my, sheet, asWb, asWs, formula, name);
    }

    // COUNTIF over a numeric column A (all > 0), open (A:A) or closed (A2:An).
    private static (MyTarget, AsTarget) BuildCountif(int n, bool open, string name)
    {
        var (my, sheet) = NewMy();
        var (asWb, asWs) = NewAspose();
        var last = n + 1;

        for (var r = 2; r <= last; r++)
        {
            double v = r;
            sheet[$"A{r}"] = new NumberValue(v);
            asWs.Cells[$"A{r}"].PutValue(v);
        }

        var range = open ? "A:A" : $"A2:A{last}";
        var formula = $"=COUNTIF({range},\">0\")";
        return Finish(my, sheet, asWb, asWs, formula, name);
    }

    // The K1 array idiom: first "Show" row strictly beyond row 2. Column A alternates Show/Hide.
    private static (MyTarget, AsTarget) BuildArray(int n, string name)
    {
        var (my, sheet) = NewMy();
        var (asWb, asWs) = NewAspose();
        var last = n + 1;

        for (var r = 2; r <= last; r++)
        {
            var s = (r % 2 == 0) ? "Show" : "Hide";
            sheet[$"A{r}"] = new StringValue(s);
            asWs.Cells[$"A{r}"].PutValue(s);
        }

        var formula =
            $"=INDEX(ROW($A:$A),SMALL(IF(A2:A{last}=\"Show\",IF(ROW(A2:A{last})>2,ROW(A2:A{last}))),1))";

        // MySheet auto-detects the array context; Aspose needs the CSE analogue via SetArrayFormula.
        var expr = ExpressionParser.Parse(formula, sheet);
        sheet[Target] = expr;
        asWs.Cells[Target].SetArrayFormula(formula, 1, 1);

        var myTarget = new MyTarget { Workbook = my, Sheet = SheetName, Cell = Target, Expression = expr };
        var asTarget = new AsTarget { Workbook = asWb, Cell = asWs.Cells[Target] };
        myTarget.Warm();
        AssertEquivalent(name, myTarget, asTarget);
        return (myTarget, asTarget);
    }

    // XLOOKUP of a key that exists mid-column, over open columns B:B / C:C.
    private static (MyTarget, AsTarget) BuildXlookup(int n, string name)
    {
        var (my, sheet) = NewMy();
        var (asWb, asWs) = NewAspose();
        var last = n + 1;

        for (var r = 2; r <= last; r++)
        {
            var key = "K" + (r - 1);
            double value = r * 10;
            sheet[$"B{r}"] = new StringValue(key);
            sheet[$"C{r}"] = new NumberValue(value);
            asWs.Cells[$"B{r}"].PutValue(key);
            asWs.Cells[$"C{r}"].PutValue(value);
        }

        var mid = "K" + (n / 2);
        var formula = $"=XLOOKUP(\"{mid}\",B:B,C:C)";
        return Finish(my, sheet, asWb, asWs, formula, name);
    }

    // MATCH + INT pagination: page (25 rows/page) of a key found mid-column.
    private static (MyTarget, AsTarget) BuildMatchInt(int n, string name)
    {
        var (my, sheet) = NewMy();
        var (asWb, asWs) = NewAspose();
        var last = n + 1;

        for (var r = 2; r <= last; r++)
        {
            var key = "K" + (r - 1);
            sheet[$"A{r}"] = new StringValue(key);
            asWs.Cells[$"A{r}"].PutValue(key);
        }

        var mid = "K" + (n / 2);
        var formula = $"=INT((MATCH(\"{mid}\",A2:A{last},0)-1)/25)+1";
        return Finish(my, sheet, asWb, asWs, formula, name);
    }

    // IF(A1="","",A1) with A1 a number — the non-empty passthrough branch.
    private static (MyTarget, AsTarget) BuildIfPassthrough(string name)
    {
        var (my, sheet) = NewMy();
        var (asWb, asWs) = NewAspose();
        sheet["A1"] = new NumberValue(42);
        asWs.Cells["A1"].PutValue(42d);
        return Finish(my, sheet, asWb, asWs, "=IF(A1=\"\",\"\",A1)", name);
    }

    // OR over 24 disjuncts; A1 matches the last one so all are evaluated.
    private static (MyTarget, AsTarget) BuildOrDisjuncts(string name)
    {
        var (my, sheet) = NewMy();
        var (asWb, asWs) = NewAspose();
        sheet["A1"] = new NumberValue(24);
        asWs.Cells["A1"].PutValue(24d);

        var terms = string.Join(",", Enumerable.Range(1, 24).Select(i => $"A1={i}"));
        return Finish(my, sheet, asWb, asWs, $"=OR({terms})", name);
    }

    private static (MyTarget, AsTarget) BuildTextUpperTrim(string name)
    {
        var (my, sheet) = NewMy();
        var (asWb, asWs) = NewAspose();
        const string v = "  hello world  ";
        sheet["A1"] = new StringValue(v);
        asWs.Cells["A1"].PutValue(v);
        return Finish(my, sheet, asWb, asWs, "=UPPER(TRIM(A1))", name);
    }

    private static (MyTarget, AsTarget) BuildTextLeftFind(string name)
    {
        var (my, sheet) = NewMy();
        var (asWb, asWs) = NewAspose();
        const string v = "ABC-12345";
        sheet["A1"] = new StringValue(v);
        asWs.Cells["A1"].PutValue(v);
        return Finish(my, sheet, asWb, asWs, "=LEFT(A1,FIND(\"-\",A1)-1)", name);
    }

    private static (MyTarget, AsTarget) BuildTextConcat(string name)
    {
        var (my, sheet) = NewMy();
        var (asWb, asWs) = NewAspose();
        sheet["A1"] = new StringValue("foo");
        sheet["B1"] = new StringValue("bar");
        sheet["C1"] = new StringValue("baz");
        asWs.Cells["A1"].PutValue("foo");
        asWs.Cells["B1"].PutValue("bar");
        asWs.Cells["C1"].PutValue("baz");
        return Finish(my, sheet, asWb, asWs, "=A1&\"-\"&B1&\"-\"&C1", name);
    }

    // Scalar sanity anchor: A1*2+B1 — should be ~0 alloc on MySheet.
    private static (MyTarget, AsTarget) BuildScalarControl(string name)
    {
        var (my, sheet) = NewMy();
        var (asWb, asWs) = NewAspose();
        sheet["A1"] = new NumberValue(3);
        sheet["B1"] = new NumberValue(4);
        asWs.Cells["A1"].PutValue(3d);
        asWs.Cells["B1"].PutValue(4d);
        return Finish(my, sheet, asWb, asWs, "=A1*2+B1", name);
    }

    // ==================================================================================================
    // Plumbing
    // ==================================================================================================

    private static (Workbook, Sheet) NewMy()
    {
        var wb = new Workbook();
        var sheet = wb.Sheets.Add(SheetName);
        return (wb, sheet);
    }

    private static (AsposeWorkbook, Worksheet) NewAspose()
    {
        var wb = new AsposeWorkbook();
        // Force every CalculateFormula() to recompute from scratch — no cached-chain shortcut (see class note).
        wb.Settings.FormulaSettings.ForceFullCalculation = true;
        wb.Settings.FormulaSettings.EnableCalculationChain = false;
        var ws = wb.Worksheets[0];
        ws.Name = SheetName;
        return (wb, ws);
    }

    // Wire the (normal) target formula on both engines, assert equivalence, return the pair.
    private static (MyTarget, AsTarget) Finish(
        Workbook my, Sheet sheet, AsposeWorkbook asWb, Worksheet asWs, string formula, string name)
    {
        var expr = ExpressionParser.Parse(formula, sheet);
        sheet[Target] = expr;
        asWs.Cells[Target].Formula = formula;

        var myTarget = new MyTarget { Workbook = my, Sheet = SheetName, Cell = Target, Expression = expr };
        var asTarget = new AsTarget { Workbook = asWb, Cell = asWs.Cells[Target] };
        myTarget.Warm();
        AssertEquivalent(name, myTarget, asTarget);
        return (myTarget, asTarget);
    }

    // Guard: both engines must compute the SAME value (numbers within tolerance, everything else ordinal).
    private static void AssertEquivalent(string name, MyTarget my, AsTarget aspose)
    {
        var mine = my.Eval();
        var theirs = aspose.Eval();

        if (TryNum(mine, out var a) && TryNum(theirs, out var b))
        {
            if (Math.Abs(a - b) > 1e-6 * Math.Max(1.0, Math.Abs(b)))
            {
                throw new InvalidOperationException(
                    $"[{name}] engine disagreement: MySheet={a} vs Aspose={b}");
            }
            return;
        }

        var ms = mine?.ToString();
        var xs = theirs?.ToString();
        if (!string.Equals(ms, xs, StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                $"[{name}] engine disagreement: MySheet='{ms ?? "<null>"}' vs Aspose='{xs ?? "<null>"}'");
        }
    }

    private static bool TryNum(object? value, out double number)
    {
        switch (value)
        {
            case double d: number = d; return true;
            case float f: number = f; return true;
            case int i: number = i; return true;
            case long l: number = l; return true;
            case decimal m: number = (double)m; return true;
            default: number = 0; return false;
        }
    }

    private static void TryLoadLicense()
    {
        var path = Environment.GetEnvironmentVariable("ASPOSE_LICENSE_PATH");
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return;
        }

        try
        {
            new License().SetLicense(path);
        }
        catch
        {
            // Stay in evaluation mode — sufficient for this fully in-memory scenario.
        }
    }
}
