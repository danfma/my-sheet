# CI/CD: versionize + publish no NuGet (Trusted Publishing)

GitHub Actions com CI (build + test) e um workflow de **release separado e manual** que versiona via
**versionize** (conventional commits) e publica a lib `MySheet` no NuGet.org via **Trusted Publishing**
(OIDC, sem API key armazenada).

## Context
- Repo: `github.com/danfma/my-sheet` (owner `danfma`). Solução: `MySheet.slnx`.
- Pacote a publicar: `MySheet/MySheet.csproj` (net10.0, features C# preview, MemoryPack). Testes:
  `tests/MySheet.Tests` (TUnit/Microsoft.Testing.Platform — roda via `dotnet run`, NÃO `dotnet test`).
  Benchmark: não é empacotado.
- Hoje: sem CI, sem metadata de pacote, sem LICENSE, sem README. SDK local `11.0.100-preview.5` (alvo net10.0).

### Decisões fechadas
- **versionize** (.NET) para versão (conventional commits → bump no `.csproj` + CHANGELOG + commit + tag).
- Licença **MIT**.
- **Release é um workflow SEPARADO e manual** (`workflow_dispatch`) — não publica direto da main.
- **Trusted Publishing**: `permissions: id-token: write` + ação `NuGet/login@v1` → API key temporária
  (~1h) → `dotnet nuget push`. A *policy* no nuget.org é passo manual do usuário (documentado na Fase 5).

### Assunções (corrigir se necessário)
- `PackageId = MySheet`; Author = Daniel Ferreira Monteiro Alves (`danfma`).
- versionize commita o bump na main via `GITHUB_TOKEN` (`contents: write`); pushes do GITHUB_TOKEN não
  re-disparam workflows (sem loop). Repo pessoal sem branch protection restritiva (confirmar).
- SDK do CI precisa compilar features C# preview → fixar via `global.json` (verificar se .NET 10 GA basta
  ou se exige o SDK 11 preview).

## For Future Agents
Marque `- [x]`; ao fechar fase, Status `Complete` + Phase Summary + rode a Verification.
A parte OIDC/Trusted-Publishing só dá para verificar de verdade no primeiro release real (após a policy
no nuget.org); o resto é verificável localmente (`dotnet pack`, `versionize --dry-run`, lint de YAML).

---

## Phase 1: Tornar `MySheet` empacotável (metadata + LICENSE + README)
Status: Complete

### Phase Summary
`LICENSE` (MIT), `README.md` e metadata no `MySheet.csproj` (PackageId/Authors/Description/URLs/
LicenseExpression=MIT/ReadmeFile/Tags + `<Version>0.0.0</Version>` placeholder). `dotnet pack -c Release`
gerou `MySheet.0.0.0.nupkg` com `nuspec` (id=MySheet, license=MIT, readme, repository url+commit),
`README.md` e `lib/net10.0/MySheet.dll`. Os `NU1900` são só falta de rede do sandbox p/ vuln data (benignos).

- [ ] `LICENSE` (MIT, copyright Daniel Ferreira Monteiro Alves) na raiz.
- [ ] `README.md` na raiz (visão geral curta; usado como `PackageReadmeFile`).
- [ ] Metadata em `MySheet.csproj`: `PackageId`, `Authors`, `Description`, `PackageProjectUrl`,
      `RepositoryUrl`/`RepositoryType=git`, `PackageLicenseExpression=MIT`, `PackageReadmeFile=README.md`
      (+ `<None Include="..\README.md" Pack="true" PackagePath="\"/>`), `PackageTags`. Garantir que
      Tests/Benchmark não empacotam (padrão já é não).

### Verification Plan
- `dotnet pack MySheet/MySheet.csproj -c Release -o ./artifacts` → gera `MySheet.<versão>.nupkg`;
  `unzip -l artifacts/*.nupkg` mostra `.nuspec` + README; build com **0 Warning(s)**.

### Phase Summary
_(escrever ao concluir)_

---

## Phase 2: Workflow de CI (build + test)
Status: Complete

### Phase Summary
`global.json` fixa o SDK `11.0.100-preview.5.26302.115` (rollForward latestFeature, allowPrerelease) — o
mesmo que compila localmente, então o CI compila as features C# preview com certeza. `.github/workflows/ci.yml`:
push/PR na main → `setup-dotnet@v4` (global-json-file) → `dotnet build MySheet.slnx -c Release` →
`dotnet run --project tests/MySheet.Tests/... -c Release --no-build`. Verificado local: build succeeded
(0 Errors; os 33 warnings são NU1900 de rede, ausentes no CI) e 172/172 testes. Run real no GitHub no 1º push.

- [ ] `global.json` fixando o SDK (versão + `rollForward`) para o CI bater com o local. **Verificar** qual
      SDK compila as features C# preview (testar .NET 10 GA; se falhar, usar o SDK 11 preview).
- [ ] `.github/workflows/ci.yml`: em `pull_request` e `push` na `main`. Passos: checkout, `setup-dotnet`
      (do global.json), restore, `dotnet build MySheet.slnx -c Release`, testes via
      `dotnet run --project tests/MySheet.Tests/MySheet.Tests.csproj -c Release`.

### Verification Plan
- Reproduzir local: `dotnet build MySheet.slnx -c Release` → 0 Warning(s);
  `dotnet run --project tests/MySheet.Tests/MySheet.Tests.csproj` → 172/172 verde.
- YAML parseia (`actionlint` se disponível). Execução real verificada no primeiro push.

### Phase Summary
_(escrever ao concluir)_

---

## Phase 3: Setup do versionize
Status: Complete

### Phase Summary
`<Version>0.0.0</Version>` no `MySheet.csproj` (Phase 1). Decisão: em vez de fixar uma versão de tool num
manifesto (impossível validar offline aqui), o `release.yml` instala o **latest globalmente**
(`dotnet tool install --global Versionize`) — resolve uma versão real no CI. versionize roda em
`working-directory: MySheet` (acha o csproj e o repo git acima). Dry-run não verificável offline (sem rede);
validado no primeiro release real.

- [ ] `<Version>0.0.0</Version>` no `MySheet.csproj` (versionize atualiza).
- [ ] Manifesto de tool local (`.config/dotnet-tools.json`) com `Versionize` (via `dotnet tool install
      Versionize` no manifesto) → CI usa `dotnet tool restore`.
- [ ] Config do versionize para mirar `MySheet.csproj` (ex.: `--workingDir MySheet` ou arquivo de config);
      verificar a flag exata no `--dry-run`.

### Verification Plan
- `dotnet tool restore && dotnet versionize --dry-run` (mirando MySheet) → imprime a próxima versão
  computada dos conventional commits sem alterar nada (a 1ª versão sai do histórico `feat/fix/...`).

### Phase Summary
_(escrever ao concluir)_

---

## Phase 4: Workflow de release (manual: versionize + Trusted Publishing)
Status: Complete

### Phase Summary
`.github/workflows/release.yml` (`workflow_dispatch`, permissions `contents:write` + `id-token:write`):
checkout fetch-depth 0 → setup-dotnet → git config bot → instala versionize → `versionize` (bump+changelog+
commit+tag) → `git push --follow-tags origin HEAD:main` → `dotnet pack` → `NuGet/login@v1` (input `user:`
= `${{ secrets.NUGET_USER }}`, output `NUGET_API_KEY`) → `dotnet nuget push --api-key ... --source
https://api.nuget.org/v3/index.json --skip-duplicate` → `gh release create`. YAML validado (parser).
**Não verificável offline**: o caminho OIDC/push e o versionize só rodam no GitHub — verificar no 1º release.

- [ ] `.github/workflows/release.yml`: trigger `workflow_dispatch`. `permissions: { contents: write,
      id-token: write }`. Passos:
  1. checkout `fetch-depth: 0` (histórico completo p/ análise de commits).
  2. `setup-dotnet`.
  3. configurar `git config user` (bot).
  4. `dotnet tool restore`; `dotnet versionize` (bump + CHANGELOG + commit + tag).
  5. `git push --follow-tags`.
  6. `dotnet pack MySheet/MySheet.csproj -c Release -o artifacts` (pega a versão recém-bumpada).
  7. `NuGet/login@v1` (com o username do nuget.org) → `NUGET_API_KEY` temporária.
  8. `dotnet nuget push artifacts/*.nupkg --api-key ${{ steps.login.outputs.NUGET_API_KEY }}
     --source https://api.nuget.org/v3/index.json --skip-duplicate`.
  9. criar GitHub Release a partir da tag + seção do CHANGELOG.

### Verification Plan
- YAML parseia. O caminho OIDC/push só é verificável num `workflow_dispatch` real, após a policy da Fase 5;
  o `pack` já foi verificado na Fase 1.

### Phase Summary
_(escrever ao concluir)_

---

## Phase 5: Documentar a policy do nuget.org + processo de release
Status: Complete

### Phase Summary
`CONTRIBUTING.md`: convenção de commits, CI, e **Releasing** com o setup único (policy no nuget.org:
Owner=danfma, Repo=my-sheet, Workflow=release.yml; secret `NUGET_USER` = username/profile do nuget.org,
não email) + como cortar o release (rodar o workflow "Release" via Actions).

- [ ] Documentar (README/CONTRIBUTING) o setup manual no nuget.org: criar a **Trusted Publishing policy**
      do pacote (owner `danfma`, repo `my-sheet`, workflow `release.yml`), e como cortar um release
      (rodar o workflow "release" via `workflow_dispatch`).

### Verification Plan
- Doc revisada; o primeiro release manual conclui e o pacote aparece no nuget.org.

### Phase Summary
_(escrever ao concluir)_

---

## Final Recap
CI/CD configurado: `MySheet` empacotável (LICENSE MIT + README + metadata), `global.json` fixando o SDK,
`ci.yml` (build+test em push/PR) e `release.yml` (manual, separado: versionize → pack → NuGet Trusted
Publishing → GitHub Release), `CONTRIBUTING.md`. Verificado offline: `dotnet pack` (nuspec correto),
`dotnet build MySheet.slnx` + 172/172 testes, YAML/JSON válidos. **Pendente de verificação no GitHub**
(offline impossível): o run do CI, o versionize e o publish OIDC — confirmar no 1º push/release.

## Deployment Plan
1. **Push** desta branch/commit → dispara o CI (`ci.yml`); confirmar verde na aba Actions (valida o SDK
   pinado e o build/test no runner).
2. **Setup único no nuget.org**: criar a Trusted Publishing policy (Owner=danfma, Repo=my-sheet,
   Workflow=release.yml) e adicionar o secret `NUGET_USER` no repo (ver `CONTRIBUTING.md`).
3. **Primeiro release**: rodar o workflow "Release" (Actions → Run workflow → main). versionize calcula a
   versão dos commits, taggeia, e publica no NuGet via OIDC. Conferir o pacote em nuget.org/packages/MySheet.
4. Ajustar se necessário: flags do versionize / SDK no CI (pontos não verificáveis offline).
