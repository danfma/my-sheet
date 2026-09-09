using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Pins the registry's array-lifting classification: the flag's default-deny zero value, the factories that
/// set it, and the exclusions that no automatic signal catches. The classification is what the mini-CSE
/// consults before lifting a function over an array element by element, so a wrong flag on a range-aware
/// function is a SILENT wrong number: it would see only the top-left element of the rectangle it was
/// supposed to consume whole.
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

    // Entry<T> is the default-deny factory: everything it builds consumes its arguments as they are.
    [Test]
    public async Task TheEntryFactory_ClassifiesAsConsumes()
    {
        await Assert.That(Of("SUM").Lifting).IsEqualTo(ArrayLifting.Consumes);
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
}
