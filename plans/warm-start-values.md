# Warm-start: persistir valores computados no Save (opt-in)

Hoje o contrato é "arquivo = modelo; valores sempre recomputados" (provado byte-idêntico frio×quente em
2026-07-03). Feature: **opt-in** para persistir a memoização junto do modelo, permitindo que um `Load`
pule a recomputação. Default de `Save`/`Load` INALTERADO (bytes idênticos aos atuais). Release minor.

## Decisões de design (defaults meus; usuário pode vetar)
- **API**: `Save(path)` inalterado (bytes atuais). Novo overload `Save(path, WorkbookSaveOptions)` com
  `IncludeComputedValues = true`; `SaveAsync` idem. `Load`/`LoadAsync` detectam o formato pelo cabeçalho.
- **Formato container, sem tocar o schema do Workbook**: arquivo warm = magic próprio (ex.: `MSWM` + byte
  de versão) + bloco do modelo (MemoryPack do Workbook, os MESMOS bytes de hoje) + bloco de valores.
  Arquivo frio = raw MemoryPack como hoje (sem magic) — `Load` diferencia pelo prefixo. Compat total:
  arquivos antigos carregam; arquivos frios novos são byte-idênticos aos antigos (o teste de hoje vira
  regressão permanente).
- **Bloco de valores**: lista de `(sheetName, cellId, valor)` com o valor num surrogate serializável do
  `ComputedValue` (Kind + double + string?/código de Error; Kind=Reference NÃO é persistido — raro como
  valor final e reconstrutível: célula fica fora do warm block e recomputa).
- **Voláteis FICAM DE FORA**: entradas em `_volatileTainted` não são persistidas (NOW()/RAND() de ontem
  não devem "descongelar" — recomputam na 1ª leitura pós-load, que re-amostra a época). Elegante: o set
  já existe.
- **Consistência por construção**: modelo e valores viajam no MESMO arquivo → nunca dessincronizam no
  load. Pós-load, o contrato é o de sempre: editou → `InvalidateCache()`.
- **Round-trip do warm**: `Load` de arquivo warm popula `_cache` direto (structs inline). Índice
  estrutural/range caches NÃO são persistidos (reconstroem sob demanda; admissão cuida do custo).

## For Future Agents
TDD; verificação `--no-incremental` 0 warnings; fixture `workbook-pre-namespaces.msgpack.bin` e
`MemoryPackCompatibilityTests` intocáveis (arquivo antigo = frio = carrega igual). Suítes hoje: core
**821**, Excel **24**. Commits inglês, semantic, SEM atribuição a IA, SEM amend. NÃO push.

## Phase 1: Container + save/load warm
Status: Not started
- [ ] `WorkbookSaveOptions { bool IncludeComputedValues }`; overloads de `Save`/`SaveAsync`.
- [ ] Container writer/reader (magic `MSWM`+versão; bloco modelo = bytes MemoryPack atuais; bloco valores
      = MemoryPack de `List<CachedCellValue>` surrogate). `Load`/`LoadAsync` sniffam o prefixo.
- [ ] Exclusão de voláteis e de Kind=Reference no snapshot do cache; Blank/Number/Boolean/Text/Error
      cobertos no surrogate.
- [ ] Testes: frio novo == frio antigo (bytes); warm round-trip (valores idênticos SEM recomputar —
      custom function com contador prova zero avaliação nos hits); volátil recomputa pós-load; arquivo
      LEGADO (fixture) carrega; warm + edição + InvalidateCache recomputa; workbook sem cache → warm ==
      frio + bloco vazio.
### Verification Plan
- Suítes verdes + novos; fixture verde; teste de bytes frio==atual.

## Phase 2: Docs + fechamento
Status: Not started
- [ ] `docs/serialization.md` (tabela persisted/not-persisted ganha a coluna warm; contrato de staleness)
      via skill `code-documentation-doc-generate`; README bullet. NÃO tocar `docs/pt-BR/`.
- [ ] Plano: fases Complete + Final Recap.
### Verification Plan
- Suítes verdes; build 0 warnings.

## Final Recap
_(ao concluir)_
## Deployment Plan
_(verificação minha → merge → push → release minor → refresh pt-BR)_
