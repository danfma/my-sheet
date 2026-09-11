using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Text;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Pins the registry's array-lifting classification: the flag's default-deny zero value, the two sibling
/// factories that set it, the 180/130 split over the 310 registered built-ins, the ROSTER that pins the
/// Elementwise half by name (see <c>TheElementwiseRoster</c> — a count alone is not a pin), and the
/// exclusions that no automatic signal catches. The classification is what the mini-CSE consults before
/// lifting a function over an array element by element, so a wrong flag on a range-aware function is a SILENT
/// wrong number: it would see only the top-left element of the rectangle it was supposed to consume whole.
/// </summary>
public class FunctionRegistryClassificationTests
{
    private static FunctionRegistry.RegistryEntry Of(string name) => FunctionRegistry.ByName[name];

    // Safe-by-default: Consumes must be the enum's ZERO value so a default-constructed entry (and any future
    // construction path that forgets the flag) means "never lift", which is today's behaviour.
    [Test]
    public async Task Consumes_IsTheZeroValue_SoADefaultEntryNeverLifts()
    {
        await Assert.That(Enum.GetValues<ArrayLifting>()[0]).IsEqualTo(ArrayLifting.Consumes);
        await Assert
            .That(default(FunctionRegistry.RegistryEntry).Lifting)
            .IsEqualTo(ArrayLifting.Consumes);
    }

    // The two factories differ in the flag and in nothing else: an Elementwise entry still carries its name,
    // arity, node type, factory and argument accessor, so it stays "one line" in the table.
    [Test]
    public async Task Entry_ClassifiesAsConsumes_AndElementwise_ClassifiesAsElementwise()
    {
        await Assert.That(Of("SUM").Lifting).IsEqualTo(ArrayLifting.Consumes);
        await Assert.That(Of("LEN").Lifting).IsEqualTo(ArrayLifting.Elementwise);
    }

    [Test]
    public async Task AnElementwiseEntry_KeepsEverythingElseTheEntryFactoryGivesIt()
    {
        var entry = Of("LEN");
        Expression[] arguments = [new Danfma.MySheet.Expressions.StringValue("abc")];
        var node = entry.Create(arguments);

        await Assert.That(entry.Name).IsEqualTo("LEN");
        await Assert.That(entry.MinArgs).IsEqualTo(1);
        await Assert.That(entry.MaxArgs).IsEqualTo(1);
        await Assert.That(entry.NodeType).IsEqualTo(typeof(Len));
        await Assert.That(node).IsTypeOf<Len>();
        await Assert.That(entry.GetArguments((Function)node)).IsEquivalentTo(arguments);
        await Assert
            .That(FunctionRegistry.ByType[typeof(Len)].Lifting)
            .IsEqualTo(ArrayLifting.Elementwise);
    }

    // The split is a derived number, not a taste: 180 pure-scalar built-ins may be lifted, the remaining 130
    // consume ranges/arrays themselves. These totals are the cheap sanity check on the shape of the table;
    // the actual pin is TheElementwiseSet_IsExactlyTheCommittedRoster below, because a total is satisfiable
    // by a compensating swap while a name is not.
    [Test]
    public async Task TheClassificationSplitsThe310BuiltInsInto180Elementwise_And130Consumes()
    {
        var entries = FunctionRegistry.ByName.Values;

        await Assert.That(entries.Length).IsEqualTo(310);
        await Assert.That(entries.Count(e => e.Lifting is ArrayLifting.Elementwise)).IsEqualTo(180);
        await Assert.That(entries.Count(e => e.Lifting is ArrayLifting.Consumes)).IsEqualTo(130);
    }

    // THE ROSTER. The exact set of names the registry flags Elementwise, committed as a sorted list, because
    // a COUNT is not a pin: the 180/130 totals above survive a compensating swap (one entry mis-flagged
    // Elementwise while another is corrected to Consumes), and they survive the honest-looking edit a
    // contributor adding a function makes — flip the factory, bump the number. Measured: mis-flagging
    // Entry<HLookup> as Elementwise<HLookup> and bumping 180→181 / 130→129 (plus the two anti-vacuity
    // constants in ElementwiseLiftingTests) left the whole suite GREEN before this roster existed, shipping a
    // range-aware lookup as liftable — a SILENT wrong number, since the lift would hand HLOOKUP one element
    // of the table it was meant to search whole.
    //
    // So the roster is what a new name has to pass through. Adding one fails this test naming the newcomer;
    // removing one fails naming the loss. Editing the roster is then a one-line diff a reviewer can see and
    // ask about — which is the whole point, because the flag itself is not derivable from the signature.
    //
    // HOW THE CLASSIFICATION WAS DERIVED, and where it is NOT oracle-verified. Each name below was measured
    // against the designated P0 oracle, Aspose.Cells 26.6.0, by handing it a rectangle and comparing the
    // array-entered answer against the element-by-element one — EXCEPT for these SIXTEEN, which Aspose does
    // not implement at all (it answers #NAME? for both the scalar and the array-entered call, measured
    // 2026-09-09): ACOT, ACOTH, ARABIC, BASE, COMBINA, CSC, CSCH, DECIMAL, FLOOR.PRECISE, ISO.CEILING,
    // PDURATION, PERMUTATIONA, PHI, RRI, SEC, SECH. Their Elementwise flag is INFERRED from the node bodies
    // (each is a closed-form scalar computation that never reads a range) and from the always-on
    // ElementwiseLiftingTests sweep, NOT measured on the oracle. A future reader must not take those sixteen
    // as oracle-verified; if Aspose ever gains them, measure them and say so here.
    private static readonly string[] TheElementwiseRoster =
    [
        "ABS",
        "ACCRINT",
        "ACCRINTM",
        "ACOS",
        "ACOSH",
        "ACOT",
        "ACOTH",
        "ADDRESS",
        "AMORDEGRC",
        "AMORLINC",
        "ARABIC",
        "ASIN",
        "ASINH",
        "ATAN",
        "ATAN2",
        "ATANH",
        "BASE",
        "CEILING",
        "CEILING.MATH",
        "CEILING.PRECISE",
        "CHAR",
        "CLEAN",
        "CODE",
        "COMBIN",
        "COMBINA",
        "COS",
        "COSH",
        "COT",
        "COTH",
        "COUPDAYBS",
        "COUPDAYS",
        "COUPDAYSNC",
        "COUPNCD",
        "COUPNUM",
        "COUPPCD",
        "CSC",
        "CSCH",
        "CUMIPMT",
        "CUMPRINC",
        "DATE",
        "DATEDIF",
        "DATEVALUE",
        "DAY",
        "DAYS",
        "DAYS360",
        "DB",
        "DDB",
        "DECIMAL",
        "DEGREES",
        "DISC",
        "DOLLAR",
        "DOLLARDE",
        "DOLLARFR",
        "DURATION",
        "EDATE",
        "EFFECT",
        "EOMONTH",
        "ERROR.TYPE",
        "EVEN",
        "EXACT",
        "EXP",
        "FACT",
        "FACTDOUBLE",
        "FIND",
        "FISHER",
        "FISHERINV",
        "FIXED",
        "FLOOR",
        "FLOOR.MATH",
        "FLOOR.PRECISE",
        "FV",
        "HOUR",
        "IFERROR",
        "IFNA",
        "IFS",
        "INT",
        "INTRATE",
        "IPMT",
        "ISBLANK",
        "ISERR",
        "ISERROR",
        "ISEVEN",
        "ISLOGICAL",
        "ISNA",
        "ISNONTEXT",
        "ISNUMBER",
        "ISO.CEILING",
        "ISODD",
        "ISOWEEKNUM",
        "ISPMT",
        "ISTEXT",
        "LEFT",
        "LEN",
        "LN",
        "LOG",
        "LOG10",
        "LOWER",
        "MDURATION",
        "MID",
        "MINUTE",
        "MOD",
        "MONTH",
        "MROUND",
        "N",
        "NOMINAL",
        "NOT",
        "NPER",
        "NUMBERVALUE",
        "ODD",
        "ODDFPRICE",
        "ODDFYIELD",
        "ODDLPRICE",
        "ODDLYIELD",
        "PDURATION",
        "PERMUT",
        "PERMUTATIONA",
        "PHI",
        "PMT",
        "POWER",
        "PPMT",
        "PRICE",
        "PRICEDISC",
        "PRICEMAT",
        "PROPER",
        "PV",
        "QUOTIENT",
        "RADIANS",
        "RATE",
        "RECEIVED",
        "REGEXEXTRACT",
        "REGEXREPLACE",
        "REGEXTEST",
        "REPLACE",
        "REPT",
        "RIGHT",
        "ROMAN",
        "ROUND",
        "ROUNDDOWN",
        "ROUNDUP",
        "RRI",
        "SEARCH",
        "SEC",
        "SECH",
        "SECOND",
        "SIGN",
        "SIN",
        "SINH",
        "SLN",
        "SQRT",
        "SQRTPI",
        "STANDARDIZE",
        "SUBSTITUTE",
        "SWITCH",
        "SYD",
        "T",
        "TAN",
        "TANH",
        "TBILLEQ",
        "TBILLPRICE",
        "TBILLYIELD",
        "TEXT",
        "TEXTAFTER",
        "TEXTBEFORE",
        "TIME",
        "TIMEVALUE",
        "TRIM",
        "TRUNC",
        "UNICHAR",
        "UNICODE",
        "UPPER",
        "VALUE",
        "VALUETOTEXT",
        "VDB",
        "WEEKDAY",
        "WEEKNUM",
        "YEAR",
        "YEARFRAC",
        "YIELD",
        "YIELDDISC",
        "YIELDMAT",
    ];

    [Test]
    public async Task TheElementwiseSet_IsExactlyTheCommittedRoster()
    {
        var actual = FunctionRegistry
            .ByName.Values.Where(e => e.Lifting is ArrayLifting.Elementwise)
            .Select(e => e.Name)
            .Order(StringComparer.Ordinal)
            .ToArray();

        // Named in BOTH directions, so the failure says which flag moved and which way.
        var difference = actual
            .Except(TheElementwiseRoster, StringComparer.Ordinal)
            .Select(name =>
                $"+{name} is now Elementwise but is not in TheElementwiseRoster — measure it on the "
                + "oracle, then add it to the roster"
            )
            .Concat(
                TheElementwiseRoster
                    .Except(actual, StringComparer.Ordinal)
                    .Select(name =>
                        $"-{name} is in TheElementwiseRoster but is no longer Elementwise — remove it "
                        + "from the roster"
                    )
            )
            .Order(StringComparer.Ordinal)
            .ToArray();

        await Assert.That(difference).IsEmpty();

        // The roster is a hand-maintained list, so keep it sorted and duplicate-free — that is what makes a
        // one-line addition to it readable in a diff.
        await Assert
            .That(string.Join(',', TheElementwiseRoster))
            .IsEqualTo(
                string.Join(
                    ',',
                    TheElementwiseRoster
                        .Order(StringComparer.Ordinal)
                        .Distinct(StringComparer.Ordinal)
                )
            );
    }

    // Neither the body-marker grep nor the executable range-awareness oracle flags these as consumers: the
    // two-population statistics and the paired-array sums errored identically on both of the oracle's
    // rectangles, PERCENTILE.EXC/TRIMMEAN answered #NUM! on both, and IF/RANDBETWEEN are excluded on design
    // grounds (IF owns a dedicated operand arm; lifting a volatile would draw once per element). Named here
    // so a future contributor cannot "fix" them by trusting the signals.
    [Test]
    [Arguments("CORREL")]
    [Arguments("COVARIANCE.P")]
    [Arguments("COVARIANCE.S")]
    [Arguments("PEARSON")]
    [Arguments("RSQ")]
    [Arguments("SLOPE")]
    [Arguments("STEYX")]
    [Arguments("FORECAST.LINEAR")]
    [Arguments("COVAR")]
    [Arguments("FORECAST")]
    [Arguments("PERCENTILE.EXC")]
    [Arguments("TRIMMEAN")]
    [Arguments("MAXIFS")]
    [Arguments("MINIFS")]
    [Arguments("SUMX2MY2")]
    [Arguments("SUMX2PY2")]
    [Arguments("SUMXMY2")]
    [Arguments("IF")]
    [Arguments("RANDBETWEEN")]
    public async Task TheHandAddedExclusions_StayConsumes(string name)
    {
        await Assert.That(Of(name).Lifting).IsEqualTo(ArrayLifting.Consumes);
    }

    // A zero-argument function can never lift (there is nothing to walk), so the flag is moot for these
    // eight; Consumes is the value that states it.
    [Test]
    public async Task TheZeroArgumentEntries_AreExactlyTheEightVolatilesAndConstants_AndAllConsume()
    {
        var zeroArgument = FunctionRegistry
            .ByName.Values.Where(e => e.MaxArgs is 0)
            .Select(e => e.Name)
            .Order()
            .ToArray();

        await Assert
            .That(zeroArgument)
            .IsEquivalentTo(
                (string[])["FALSE", "NA", "NOW", "PI", "RAND", "SHEETS", "TODAY", "TRUE"]
            );
        await Assert.That(zeroArgument.All(n => Of(n).Lifting is ArrayLifting.Consumes)).IsTrue();
    }

    // Every liftable entry must have somewhere to put an array; the converse of the zero-argument rule.
    [Test]
    public async Task EveryElementwiseEntry_TakesAtLeastOneArgument()
    {
        var argumentless = FunctionRegistry
            .ByName.Values.Where(e => e.Lifting is ArrayLifting.Elementwise && e.MaxArgs is 0)
            .Select(e => e.Name)
            .ToArray();

        await Assert.That(argumentless).IsEmpty();
    }

    // The guard Phase 7's final review asked for. A producer flagged Elementwise used to recurse to a STACK
    // OVERFLOW — the lift arm precedes the producer arm in ArrayEvaluation's switches, so the lift asks for
    // the node's scalar, FirstElement builds the array operand, and that build re-enters the lift. A crash
    // takes the whole host down and no test can catch it, which is strictly worse than a wrong number.
    [Test]
    public async Task AProducerFlaggedElementwise_IsRefusedByName_RatherThanOverflowingTheStack()
    {
        // Every producer in the tree, fabricated with the WRONG flag. The real entries are Consumes, so this
        // reaches the validator the only way a test can: by building the entry the mistake would build.
        foreach (var (name, nodeType) in Producers())
        {
            var misflagged = new FunctionRegistry.RegistryEntry(
                name,
                1,
                1,
                static arguments => arguments[0],
                nodeType,
                static f => [],
                ArrayLifting.Elementwise
            );

            var thrown = Assert.Throws<InvalidOperationException>(() =>
                FunctionRegistry.RequireProducerIsConsumes(misflagged)
            );

            // The message must NAME the offender, or the guard is a riddle at type-initialization time.
            await Assert.That(thrown!.Message).Contains(name);
            await Assert.That(thrown.Message).Contains(nodeType.Name);
            await Assert.That(thrown.Message).Contains("Entry<T>");
        }

        // Anti-vacuity: the same validator must ACCEPT each of them with the right flag, and accept a
        // genuinely elementwise entry. Without this the test would pass if the validator threw on everything.
        foreach (var (name, nodeType) in Producers())
        {
            FunctionRegistry.RequireProducerIsConsumes(
                new FunctionRegistry.RegistryEntry(
                    name,
                    1,
                    1,
                    static arguments => arguments[0],
                    nodeType,
                    static f => [],
                    ArrayLifting.Consumes
                )
            );
        }

        // And the shipped table passes it, entry by entry — the static constructor already ran by the time
        // this test executes, but asserting it here is what makes the coverage visible.
        foreach (var entry in FunctionRegistry.ByName.Values)
        {
            FunctionRegistry.RequireProducerIsConsumes(entry);
        }

        // The roster this test walks is not hand-written: it is every node type in the assembly that
        // implements the interface, so a fifth producer is covered the day it is added.
        static (string Name, Type NodeType)[] Producers() =>
            typeof(Workbook)
                .Assembly.GetTypes()
                .Where(t =>
                    t is { IsAbstract: false, IsInterface: false }
                    && t.GetInterfaces().Any(i => i.Name == "IArrayProducer")
                )
                .Select(t => (t.Name.ToUpperInvariant(), t))
                .OrderBy(pair => pair.Item1)
                .ToArray();
    }

    // The roster above must not be empty, or the loop asserts nothing at all.
    [Test]
    public async Task TheProducerRoster_IsTheFourThisPhaseAdded()
    {
        var producers = typeof(Workbook)
            .Assembly.GetTypes()
            .Where(t =>
                t is { IsAbstract: false, IsInterface: false }
                && t.GetInterfaces().Any(i => i.Name == "IArrayProducer")
            )
            .Select(t => t.Name)
            .OrderBy(name => name)
            .ToArray();

        await Assert
            .That(producers)
            .IsEquivalentTo((string[])["Filter", "Sequence", "Sort", "Unique"]);
    }
}
