namespace Danfma.MySheet.Tests;

/// <summary>
/// A deterministic <see cref="TimeProvider"/> for the volatile-function tests: a settable UTC instant and a
/// fixed local time zone, so <c>NOW()</c>/<c>TODAY()</c> are reproducible and the LOCAL-time path is
/// exercised without depending on the machine's clock or its configured zone. This is a plain override of
/// the BCL abstraction — no NuGet package. <see cref="TimeProvider.GetLocalNow"/> derives the local instant
/// from <see cref="GetUtcNow"/> and <see cref="LocalTimeZone"/>, so overriding those two is enough.
/// </summary>
public sealed class FixedTimeProvider(DateTimeOffset utcNow, TimeZoneInfo localZone) : TimeProvider
{
    private readonly TimeZoneInfo _localZone = localZone;

    /// <summary>The current UTC instant; mutate it (or call <see cref="Advance"/>) to move the clock.</summary>
    public DateTimeOffset UtcNow { get; set; } = utcNow;

    public override DateTimeOffset GetUtcNow() => UtcNow;

    public override TimeZoneInfo LocalTimeZone => _localZone;

    /// <summary>Moves the clock forward by <paramref name="by"/>.</summary>
    public void Advance(TimeSpan by) => UtcNow += by;
}
