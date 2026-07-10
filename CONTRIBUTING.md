# Contributing

## Commits

The version is derived from [Conventional Commits](https://www.conventionalcommits.org/) by
[versionize](https://github.com/versionize/versionize). Use `feat:`, `fix:`, `perf:`, `refactor:`,
`test:`, `chore:`, etc. A `feat:` bumps the minor version, `fix:`/`perf:` the patch; a `BREAKING CHANGE:`
footer (or `!`) bumps the major.

## Git hooks

Git hooks are managed by [Husky.Net](https://alirezanet.github.io/Husky.Net/) and install
themselves — the `Directory.Build.props` target runs `dotnet husky install` on the first restore,
so a fresh clone wires up `core.hooksPath` -> `.husky` with no manual step. Formatting is enforced
with [CSharpier](https://csharpier.com/); both tools are pinned in `dotnet-tools.json`.

- **pre-commit** — `dotnet csharpier check .` + `dotnet build` (Debug).
- **pre-push** — `dotnet csharpier check .` + `dotnet build` (Release) + both test suites.

Run `dotnet csharpier format .` to fix formatting. Bypass a hook when you must with `--no-verify`
(e.g. `git commit --no-verify`) or disable all hooks for a command with `HUSKY=0`.

## CI

`.github/workflows/ci.yml` runs on every push and pull request to `main`: it builds `Danfma.MySheet.slnx`
(Release) and runs the test suite (`dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj`).
The .NET SDK is pinned in `global.json`.

## Releasing

Releases are **manual** and separate from `main` — pushing to `main` never publishes. The release runs
through `.github/workflows/release.yml` (`workflow_dispatch`), which: bumps the version + `CHANGELOG.md`
via versionize, commits and tags, packs `Danfma.MySheet`, and publishes to NuGet.org using
**Trusted Publishing** (OIDC — no stored API key), then creates a GitHub Release.

### One-time setup

1. **NuGet.org Trusted Publishing policy** — log into nuget.org → your username → **Trusted Publishing**
   → add a policy with:
   - Repository Owner: `danfma`
   - Repository: `my-sheet`
   - Workflow File: `release.yml` (file name only, no path)
   - Environment: *(leave empty)*

   The first release activates the policy permanently (private repos get a 7-day activation window).

2. **`NUGET_USER` secret** — in the GitHub repo: Settings → Secrets and variables → Actions → add a
   secret `NUGET_USER` with your **nuget.org username (profile name)** — not your email.

### Cutting a release

Run the **Release** workflow from the Actions tab (Run workflow → `main`). versionize computes the next
version from the commits since the last tag; if there are no releasable commits it makes no changes.
