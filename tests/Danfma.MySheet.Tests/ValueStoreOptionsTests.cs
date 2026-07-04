using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet.Tests;

/// <summary>
/// Tests for the configurable dense value-store geometry (plans/dense-value-store-4.0.md — Phase 2 follow-up):
/// the page/group sizes and sparsity thresholds are tunable per workbook via <see cref="WorkbookOptions"/>,
/// the defaults reproduce the shipped behavior, invalid sizes are rejected at construction, and — crucially —
/// the options are runtime CONFIG (never serialized), so a workbook restored from disk silently falls back to
/// the defaults with the wire schema untouched.
/// </summary>
public class ValueStoreOptionsTests
{
    // === Defaults preserved =============================================================================

    [Test]
    public async Task DefaultWorkbook_UsesDefaultGeometry()
    {
        var store = new Workbook().ValueStoreForTesting;

        await Assert.That(store.ConfiguredPageRows).IsEqualTo(1024);
        await Assert.That(store.ConfiguredColumnGroupSize).IsEqualTo(64);
        await Assert.That(store.ConfiguredSparsityWarmupPages).IsEqualTo(64);
        await Assert.That(store.ConfiguredSparsityMinCellsPerPage).IsEqualTo(4);
        await Assert.That(store.ConfiguredInitialPageSlots).IsEqualTo(128);
    }

    [Test]
    public async Task ExplicitDefaultOptions_MatchTheParameterlessConstructor()
    {
        var store = new Workbook(WorkbookOptions.Default).ValueStoreForTesting;

        await Assert.That(store.ConfiguredPageRows).IsEqualTo(1024);
        await Assert.That(store.ConfiguredColumnGroupSize).IsEqualTo(64);
    }

    // === Custom geometry flows through AND computes end-to-end ===========================================

    [Test]
    public async Task CustomPageSize_FlowsToStore_AndComputesCorrectly()
    {
        var workbook = new Workbook(
            new WorkbookOptions { ValueStore = new ValueStoreOptions { RowPageSize = 256 } }
        );

        // The store the workbook builds must actually use the configured page size (not silently ignore it).
        await Assert.That(workbook.ValueStoreForTesting.ConfiguredPageRows).IsEqualTo(256);

        var sheet = workbook.Sheets.Add("S");
        for (var row = 1; row <= 600; row++) // spans three 256-row pages (rows 1..255, 256..511, 512..600)
        {
            sheet["A" + row] = new NumberValue(row);
        }

        sheet["B1"] = ExpressionParser.Parse("=SUM(A1:A600)", sheet); // 600*601/2 = 180300
        sheet["C1"] = ExpressionParser.Parse("=A300*2", sheet);       // reads across a page boundary

        await Assert.That(workbook.GetCellValue("S", "B1").ToDouble()).IsEqualTo(180300.0);
        await Assert.That(workbook.GetCellValue("S", "C1").ToDouble()).IsEqualTo(600.0);

        // Re-read every populated cell: nothing lost across the smaller pages.
        var mismatches = 0;
        for (var row = 1; row <= 600; row++)
        {
            if (workbook.GetCellValue("S", "A" + row).ToDouble() != row)
            {
                mismatches++;
            }
        }

        await Assert.That(mismatches).IsEqualTo(0);
    }

    [Test]
    public async Task CustomColumnGroupSize_FlowsToStore_AndSpansGroups()
    {
        var workbook = new Workbook(
            new WorkbookOptions { ValueStore = new ValueStoreOptions { ColumnGroupSize = 16 } }
        );

        await Assert.That(workbook.ValueStoreForTesting.ConfiguredColumnGroupSize).IsEqualTo(16);

        var sheet = workbook.Sheets.Add("S");
        // Columns A..AZ (1..52) span multiple 16-column groups; store one value per column and read it back.
        for (var col = 1; col <= 52; col++)
        {
            sheet[new CellAddress(col, 1).ToId()] = new NumberValue(col * 10);
        }

        var mismatches = 0;
        for (var col = 1; col <= 52; col++)
        {
            if (workbook.GetCellValue("S", new CellAddress(col, 1).ToId()).ToDouble() != col * 10)
            {
                mismatches++;
            }
        }

        await Assert.That(mismatches).IsEqualTo(0);
    }

    [Test]
    public async Task CustomInitialPageSlots_FlowsToStore_AndPromotesOnDemand()
    {
        // A small initial page (16 slots) still covers the full 1024-row interval and grows by doubling as higher
        // rows are written — the sheet computes correctly across the promotions.
        var workbook = new Workbook(
            new WorkbookOptions { ValueStore = new ValueStoreOptions { InitialPageSlots = 16 } }
        );

        await Assert.That(workbook.ValueStoreForTesting.ConfiguredInitialPageSlots).IsEqualTo(16);

        var sheet = workbook.Sheets.Add("S");
        for (var row = 1; row <= 500; row++) // forces the page from 16 up toward 512 within a single 1024-row page
        {
            sheet["A" + row] = new NumberValue(row);
        }

        sheet["B1"] = ExpressionParser.Parse("=SUM(A1:A500)", sheet); // 500*501/2 = 125250

        await Assert.That(workbook.GetCellValue("S", "B1").ToDouble()).IsEqualTo(125250.0);

        var mismatches = 0;
        for (var row = 1; row <= 500; row++)
        {
            if (workbook.GetCellValue("S", "A" + row).ToDouble() != row)
            {
                mismatches++;
            }
        }

        await Assert.That(mismatches).IsEqualTo(0);
    }

    // === Validation: bad sizes/ranges rejected at construction ===========================================

    [Test]
    public async Task NonPowerOfTwoRowPageSize_Throws()
    {
        await Assert
            .That(() => new Workbook(Options(new ValueStoreOptions { RowPageSize = 1000 })))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task RowPageSizeBelowMinimum_Throws()
    {
        await Assert
            .That(() => new Workbook(Options(new ValueStoreOptions { RowPageSize = 32 })))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task RowPageSizeAboveMaximum_Throws()
    {
        await Assert
            .That(() => new Workbook(Options(new ValueStoreOptions { RowPageSize = 131_072 })))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task NonPowerOfTwoColumnGroupSize_Throws()
    {
        await Assert
            .That(() => new Workbook(Options(new ValueStoreOptions { ColumnGroupSize = 48 })))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task ColumnGroupSizeAboveMaximum_Throws()
    {
        await Assert
            .That(() => new Workbook(Options(new ValueStoreOptions { ColumnGroupSize = 8192 })))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task SparsityWarmupPagesBelowMinimum_Throws()
    {
        await Assert
            .That(() => new Workbook(Options(new ValueStoreOptions { SparsityWarmupPages = 0 })))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task SparsityMinCellsPerPageAboveRowPageSize_Throws()
    {
        // 2000 cells/page is impossible with a 1024-row page — rejected as out of range.
        await Assert
            .That(() => new Workbook(Options(new ValueStoreOptions { SparsityMinCellsPerPage = 2000 })))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task NonPowerOfTwoInitialPageSlots_Throws()
    {
        await Assert
            .That(() => new Workbook(Options(new ValueStoreOptions { InitialPageSlots = 100 })))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task InitialPageSlotsBelowMinimum_Throws()
    {
        // 8 is below the 16-slot floor.
        await Assert
            .That(() => new Workbook(Options(new ValueStoreOptions { InitialPageSlots = 8 })))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task InitialPageSlotsAboveRowPageSize_Throws()
    {
        // A page never grows past its row span, so InitialPageSlots > RowPageSize is out of range.
        await Assert
            .That(() =>
                new Workbook(
                    Options(new ValueStoreOptions { RowPageSize = 256, InitialPageSlots = 512 })
                )
            )
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task ValidCustomGeometry_DoesNotThrow_AndFlowsThrough()
    {
        var workbook = new Workbook(
            Options(
                new ValueStoreOptions
                {
                    RowPageSize = 512,
                    ColumnGroupSize = 32,
                    InitialPageSlots = 64,
                    SparsityWarmupPages = 8,
                    SparsityMinCellsPerPage = 2,
                }
            )
        );

        var store = workbook.ValueStoreForTesting;
        await Assert.That(store.ConfiguredPageRows).IsEqualTo(512);
        await Assert.That(store.ConfiguredColumnGroupSize).IsEqualTo(32);
        await Assert.That(store.ConfiguredInitialPageSlots).IsEqualTo(64);
        await Assert.That(store.ConfiguredSparsityWarmupPages).IsEqualTo(8);
        await Assert.That(store.ConfiguredSparsityMinCellsPerPage).IsEqualTo(2);
    }

    // === Config is runtime-only: never serialized, defaults after Load, schema unchanged =================

    [Test]
    public async Task Options_AreNotSerialized_SchemaByteIdenticalToDefault()
    {
        // Two workbooks with identical CONTENT but different store options must serialize to the SAME bytes —
        // proof that ValueStoreOptions is [MemoryPackIgnore] and never touches the wire schema.
        var custom = BuildRepresentative(
            new WorkbookOptions
            {
                ValueStore = new ValueStoreOptions { RowPageSize = 256, ColumnGroupSize = 16 },
            }
        );
        var withDefaults = BuildRepresentative(WorkbookOptions.Default);

        var customBytes = MemoryPackSerializer.Serialize(custom);
        var defaultBytes = MemoryPackSerializer.Serialize(withDefaults);

        await Assert.That(customBytes.AsSpan().SequenceEqual(defaultBytes)).IsTrue();
    }

    [Test]
    public async Task DeserializedWorkbook_FallsBackToDefaultGeometry_AndEvaluates()
    {
        var custom = BuildRepresentative(
            new WorkbookOptions
            {
                ValueStore = new ValueStoreOptions { RowPageSize = 256, ColumnGroupSize = 16 },
            }
        );

        var loaded = MemoryPackSerializer.Deserialize<Workbook>(MemoryPackSerializer.Serialize(custom))!;

        // The custom options are gone (runtime config, not persisted): the restored workbook uses the defaults,
        // with no NullReferenceException from the bypassed field initializer (the F1 lesson).
        await Assert.That(loaded.ValueStoreForTesting.ConfiguredPageRows).IsEqualTo(1024);
        await Assert.That(loaded.ValueStoreForTesting.ConfiguredColumnGroupSize).IsEqualTo(64);

        // And it still evaluates correctly on the default geometry.
        await Assert.That(loaded.GetCellValue("S", "B1").ToDouble()).IsEqualTo(20.0);
    }

    private static WorkbookOptions Options(ValueStoreOptions valueStore) =>
        new() { ValueStore = valueStore };

    private static Workbook BuildRepresentative(WorkbookOptions options)
    {
        var workbook = new Workbook(options);
        var sheet = workbook.Sheets.Add("S");
        sheet["A1"] = new NumberValue(10);
        sheet["B1"] = ExpressionParser.Parse("=A1*2", sheet);
        return workbook;
    }
}
