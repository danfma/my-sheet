# Change Log

All notable changes to this project will be documented in this file. See [versionize](https://github.com/versionize/versionize) for commit guidelines.

<a name="3.12.3"></a>
## [3.12.3](https://www.github.com/danfma/my-sheet/releases/tag/v3.12.3) (2026-07-10)

<a name="3.12.2"></a>
## [3.12.2](https://www.github.com/danfma/my-sheet/releases/tag/v3.12.2) (2026-07-10)

<a name="3.12.1"></a>
## [3.12.1](https://www.github.com/danfma/my-sheet/releases/tag/v3.12.1) (2026-07-10)

<a name="3.12.0"></a>
## [3.12.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.12.0) (2026-07-10)

### Features

* **api:** optional arity validation for custom functions (docs + release marker) ([dcf6727](https://www.github.com/danfma/my-sheet/commit/dcf672755b2b45d68eebcdd6ac53f8f159fc30b5))

<a name="3.11.0"></a>
## [3.11.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.11.0) (2026-07-10)

### Features

* **excel-load:** ExcelLoadOptions with load-warning callback ([182a1f4](https://www.github.com/danfma/my-sheet/commit/182a1f48535adc38ca6538e75e38895a2b6071d7))

<a name="3.10.4"></a>
## [3.10.4](https://www.github.com/danfma/my-sheet/releases/tag/v3.10.4) (2026-07-10)

<a name="3.10.3"></a>
## [3.10.3](https://www.github.com/danfma/my-sheet/releases/tag/v3.10.3) (2026-07-10)

<a name="3.10.2"></a>
## [3.10.2](https://www.github.com/danfma/my-sheet/releases/tag/v3.10.2) (2026-07-10)

<a name="3.10.1"></a>
## [3.10.1](https://www.github.com/danfma/my-sheet/releases/tag/v3.10.1) (2026-07-10)

### Bug Fixes

* **excel-merge:** support implicit row/cell positions in the merge target ([23abe0f](https://www.github.com/danfma/my-sheet/commit/23abe0f9613817afd2cea68c04b47d4b7c29538f))
* **parse:** bound formula nesting depth in Parser and FormulaWriter ([d22cea1](https://www.github.com/danfma/my-sheet/commit/d22cea18a636aaec238161dc81e0b39fbf8f9a13))
* **recalc:** detect defined-name redefinition as graph staleness ([ba7a29f](https://www.github.com/danfma/my-sheet/commit/ba7a29f9ae9ca612a5672ad40eb08d181111d9ed))

<a name="3.10.0"></a>
## [3.10.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.10.0) (2026-07-10)

### Features

* **api:** populated-cell enumeration and span-based id formatting ([acd2a16](https://www.github.com/danfma/my-sheet/commit/acd2a16e48737ce04d1d0c956885b6cc745d1717))

<a name="3.9.0"></a>
## [3.9.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.9.0) (2026-07-10)

### Features

* **api:** add SheetValueReader — numeric-address bulk value reads ([745ff9a](https://www.github.com/danfma/my-sheet/commit/745ff9af107028ce849ffd78593bbefd0bf560bd))

<a name="3.8.3"></a>
## [3.8.3](https://www.github.com/danfma/my-sheet/releases/tag/v3.8.3) (2026-07-09)

<a name="3.8.2"></a>
## [3.8.2](https://www.github.com/danfma/my-sheet/releases/tag/v3.8.2) (2026-07-09)

<a name="3.8.1"></a>
## [3.8.1](https://www.github.com/danfma/my-sheet/releases/tag/v3.8.1) (2026-07-09)

### Bug Fixes

* **excel-export:** preserve leading/trailing whitespace in cell text ([d3315f6](https://www.github.com/danfma/my-sheet/commit/d3315f6269d6db43a36d69117f1d120706904855))

<a name="3.8.0"></a>
## [3.8.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.8.0) (2026-07-08)

### Features

* **recalc:** incremental recomputation engine with reverse dependency graph ([2f2f4d9](https://www.github.com/danfma/my-sheet/commit/2f2f4d970e6b0ce614b30a841a43c6e2ed1afc00))

<a name="3.7.0"></a>
## [3.7.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.7.0) (2026-07-08)

### Features

* **functions:** implement INDIRECT (A1-style, volatile, works as a ':' endpoint) ([7c65c58](https://www.github.com/danfma/my-sheet/commit/7c65c58762a091e3c903811b3eccba97ad563e2f))
* **parser:** cross-sheet ':' ranges yield #REF!; sheet-qualified right endpoints parse ([0bcf3a5](https://www.github.com/danfma/my-sheet/commit/0bcf3a57132a1fb9dc1444797c3b4acc7da926d0))

### Bug Fixes

* **excel-merge:** drop stale calcChain so Excel does not force a repair ([b09290b](https://www.github.com/danfma/my-sheet/commit/b09290b88aaed93815de0c8ba8627aac9e0c0222))
* **formula-writer:** render INDIRECT so ToFormula/FORMULATEXT/SaveAsExcel round-trip ([935aed0](https://www.github.com/danfma/my-sheet/commit/935aed095002ccde2a3ef8e0cefdfff90537ca1d))
* **offset:** non-positive height/width is #REF!; report ':' endpoint error at the right side ([2f3804b](https://www.github.com/danfma/my-sheet/commit/2f3804bf4ebdaec7d6a3842086d27378165bc90b))
* **offset:** truncate height/width toward zero to match Excel ([4cf68e1](https://www.github.com/danfma/my-sheet/commit/4cf68e19e91e923694deb1b2941b405539143a1c))

<a name="3.6.0"></a>
## [3.6.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.6.0) (2026-07-07)

### Features

* **excel-merge:** add overload to skip a set of sheets during merge ([6f32abe](https://www.github.com/danfma/my-sheet/commit/6f32abe2f967a43df5f225e1c4e05d8fdff3c9fc))

<a name="3.5.0"></a>
## [3.5.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.5.0) (2026-07-07)

### Features

* **choose:** resolve CHOOSE to the chosen argument's reference ([6cfeee2](https://www.github.com/danfma/my-sheet/commit/6cfeee2196e680f9d564ee57b9f461f4b841fa1c))
* **expr:** add DynamicRange node spanning reference-expression endpoints ([c829941](https://www.github.com/danfma/my-sheet/commit/c8299414972f286c97213204076a1ee1b33b8958))
* **expr:** add TryResolveReference resolution path on Expression and references ([45b327d](https://www.github.com/danfma/my-sheet/commit/45b327d4153e368074e7f4e44d0021badc12e8a6))
* **index:** resolve INDEX to a target reference in reference context ([78f18d3](https://www.github.com/danfma/my-sheet/commit/78f18d3cc006e0810aba4eec06dc77050b4f92f7))
* **offset:** resolve OFFSET to a target reference; share compute with Evaluate ([e11fedc](https://www.github.com/danfma/my-sheet/commit/e11fedc89b93f01895f76a95d022c943f6cf1910))
* **parser:** build DynamicRange for non-static ':' endpoints instead of throwing ([1f7df5f](https://www.github.com/danfma/my-sheet/commit/1f7df5f9eb737a5481231e7f6ff881bdffb2bdc9))

### Bug Fixes

* **formula-writer:** render DynamicRange endpoints ([4d5eadf](https://www.github.com/danfma/my-sheet/commit/4d5eadf9697d4070b46061af07be914c3599c385))
* **index:** normalize range corners when resolving INDEX to a reference ([3de0336](https://www.github.com/danfma/my-sheet/commit/3de0336bcaa1d3aebff97bc5ca0c455f4b2b8dfe))
* **offset:** compare height/width as doubles to preserve original Evaluate ([aa0d2e3](https://www.github.com/danfma/my-sheet/commit/aa0d2e33169bf64d6205bbf2731ade4db0d1e312))
* **offset:** preserve specific argument errors when resolving OFFSET ([11ac0ab](https://www.github.com/danfma/my-sheet/commit/11ac0ab3699df068895be7b194d0c66adfb4aefb))
* **reference-guard:** guard DynamicRange over a missing sheet ([29995f5](https://www.github.com/danfma/my-sheet/commit/29995f51fa6abf4784ca2f96e2b52921c07b0674))

<a name="3.4.1"></a>
## [3.4.1](https://www.github.com/danfma/my-sheet/releases/tag/v3.4.1) (2026-07-06)

<a name="3.4.0"></a>
## [3.4.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.4.0) (2026-07-06)

### Features

* **save:** expose Brotli CompressionLevel on WorkbookSaveOptions ([e7421e1](https://www.github.com/danfma/my-sheet/commit/e7421e1bf8f2af74b88e7b4df9999909d63a24be))

<a name="3.3.0"></a>
## [3.3.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.3.0) (2026-07-06)

### Features

* **workbook:** add public ComputeAll for eager full evaluation ([f29e1c6](https://www.github.com/danfma/my-sheet/commit/f29e1c68dbc07e6e646a6546f3c7a180ee6c39e9))

<a name="3.2.2"></a>
## [3.2.2](https://www.github.com/danfma/my-sheet/releases/tag/v3.2.2) (2026-07-04)

<a name="3.2.1"></a>
## [3.2.1](https://www.github.com/danfma/my-sheet/releases/tag/v3.2.1) (2026-07-04)

<a name="3.2.0"></a>
## [3.2.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.2.0) (2026-07-04)

### Features

* **store:** make the dense value store geometry configurable ([4f84a2d](https://www.github.com/danfma/my-sheet/commit/4f84a2df0bc525b59606a92b7f22a7f431925a3f))

<a name="3.1.1"></a>
## [3.1.1](https://www.github.com/danfma/my-sheet/releases/tag/v3.1.1) (2026-07-04)

### Bug Fixes

* **logical:** ignore literal text/blank operands in AND/OR/XOR ([4fdfe4a](https://www.github.com/danfma/my-sheet/commit/4fdfe4a76606e1105b72baa000be4bd9af3ac4c0))

<a name="3.1.0"></a>
## [3.1.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.1.0) (2026-07-04)

### Features

* **eval:** add internal ArrayEvaluation element-wise evaluator (mini-CSE Phase A) ([2ee7deb](https://www.github.com/danfma/my-sheet/commit/2ee7deb3f380234324378ddb4b57cd4b9fc4b566))
* **eval:** wire mini-CSE consumers to the element-wise evaluator ([f66161d](https://www.github.com/danfma/my-sheet/commit/f66161d79643c769a68ad4425793604ed14c5c27))

<a name="3.0.0"></a>
## [3.0.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.0.0) (2026-07-03)

### Features

* **sheet:** encapsulate the cell store behind a write choke point ([93826a0](https://www.github.com/danfma/my-sheet/commit/93826a0ab5f140a1feba205371e2099c9dada77f))
* **sheet:** make the structural index write-maintained and lifetime-scoped ([1eb8c17](https://www.github.com/danfma/my-sheet/commit/1eb8c174fa225e68cd418927bb6e9f01a5af6435))

### Breaking Changes

* **sheet:** encapsulate the cell store behind a write choke point ([93826a0](https://www.github.com/danfma/my-sheet/commit/93826a0ab5f140a1feba205371e2099c9dada77f))

<a name="2.9.1"></a>
## [2.9.1](https://www.github.com/danfma/my-sheet/releases/tag/v2.9.1) (2026-07-03)

### Bug Fixes

* **logical:** ignore text and blank from references in OR/AND/XOR ([6a4cf7c](https://www.github.com/danfma/my-sheet/commit/6a4cf7cd1490014a5b885eb90765c6db60b7d3fd))

<a name="2.9.0"></a>
## [2.9.0](https://www.github.com/danfma/my-sheet/releases/tag/v2.9.0) (2026-07-03)

### Features

* **core:** add optional Brotli compression to workbook saves ([1b7b907](https://www.github.com/danfma/my-sheet/commit/1b7b9079305ccb7c077f328360cc1859dddf2273))

<a name="2.8.1"></a>
## [2.8.1](https://www.github.com/danfma/my-sheet/releases/tag/v2.8.1) (2026-07-03)

<a name="2.8.0"></a>
## [2.8.0](https://www.github.com/danfma/my-sheet/releases/tag/v2.8.0) (2026-07-03)

### Features

* **core:** persist computed values for warm-start loads ([8887d2e](https://www.github.com/danfma/my-sheet/commit/8887d2ebbb58af3e522f9899f7cd73bfb17ae12d))

<a name="2.7.0"></a>
## [2.7.0](https://www.github.com/danfma/my-sheet/releases/tag/v2.7.0) (2026-07-03)

### Features

* **core:** coerce empty formula results to zero at the cell boundary ([6ecf004](https://www.github.com/danfma/my-sheet/commit/6ecf00426932dc82e9b8ccd74da517725323f33d))

<a name="2.6.4"></a>
## [2.6.4](https://www.github.com/danfma/my-sheet/releases/tag/v2.6.4) (2026-07-03)

<a name="2.6.3"></a>
## [2.6.3](https://www.github.com/danfma/my-sheet/releases/tag/v2.6.3) (2026-07-03)

<a name="2.6.2"></a>
## [2.6.2](https://www.github.com/danfma/my-sheet/releases/tag/v2.6.2) (2026-07-03)

<a name="2.6.1"></a>
## [2.6.1](https://www.github.com/danfma/my-sheet/releases/tag/v2.6.1) (2026-07-03)

### Bug Fixes

* **refs:** resolve missing-sheet #REF! across the long-tail consumers ([278b079](https://www.github.com/danfma/my-sheet/commit/278b07962b4f13425ed85f7ed32ae3d314ad5c8e))
* **refs:** resolve missing-sheet reference to #REF! instead of throwing ([f903a77](https://www.github.com/danfma/my-sheet/commit/f903a773160e18af538ced2b037693f9f1c580b9))

<a name="2.6.0"></a>
## [2.6.0](https://www.github.com/danfma/my-sheet/releases/tag/v2.6.0) (2026-07-03)

### Features

* **benchmark:** add whole-column reference storage spike ([055a1bb](https://www.github.com/danfma/my-sheet/commit/055a1bb85d5ca0936bb9b90cdd7233637b8ac0df))
* **refs:** parse and aggregate whole-column/row references ([ca76882](https://www.github.com/danfma/my-sheet/commit/ca768827076324a825887b352f1cefb6b3c09c98))
* **refs:** resolve whole-column references in reference consumers ([3b5d161](https://www.github.com/danfma/my-sheet/commit/3b5d1619bebd2a1601fd7139e9cdfcc98b68de14))
* **refs:** un-parse and .xlsx interop for whole-column references ([72f932e](https://www.github.com/danfma/my-sheet/commit/72f932efc1977e4be879dfb4254546baab44b3da))

<a name="2.5.0"></a>
## [2.5.0](https://www.github.com/danfma/my-sheet/releases/tag/v2.5.0) (2026-07-02)

### Features

* **volatile:** add NOW/TODAY with an epoch cache model ([6a50953](https://www.github.com/danfma/my-sheet/commit/6a509532adea681ccc94a91e4e515e6e48b814e5))
* **volatile:** add RAND/RANDBETWEEN on a seedable RNG ([4c43881](https://www.github.com/danfma/my-sheet/commit/4c43881d663810855ffb32cb88de26f4ea012105))

<a name="2.4.0"></a>
## [2.4.0](https://www.github.com/danfma/my-sheet/releases/tag/v2.4.0) (2026-07-02)

### Features

* **functions:** add the remaining viable financial functions (wave 6) ([eb12194](https://www.github.com/danfma/my-sheet/commit/eb12194044d0d369879f41d6694817d4f6384e87))

<a name="2.3.0"></a>
## [2.3.0](https://www.github.com/danfma/my-sheet/releases/tag/v2.3.0) (2026-07-02)

### Features

* **functions:** add date and time wave ([f92fd8b](https://www.github.com/danfma/my-sheet/commit/f92fd8b4fe524c5c3910010a2db10b0e1ac316cf))

<a name="2.2.0"></a>
## [2.2.0](https://www.github.com/danfma/my-sheet/releases/tag/v2.2.0) (2026-07-02)

### Features

* **names:** add workbook-level defined names ([f3b0ca8](https://www.github.com/danfma/my-sheet/commit/f3b0ca8823999747b9d2b36191fb7928284211c9))
* **names:** read and write defined names in xlsx interop ([5e3c7da](https://www.github.com/danfma/my-sheet/commit/5e3c7dacec64cf20f46bcb1d306210f63cdb6142))

### Bug Fixes

* **excel:** guard nullable workbook part when writing defined names ([bf5838f](https://www.github.com/danfma/my-sheet/commit/bf5838fbcd39a60c799393e0749bc5e00ffa0d06))

<a name="2.1.0"></a>
## [2.1.0](https://www.github.com/danfma/my-sheet/releases/tag/v2.1.0) (2026-07-02)

### Features

* **functions:** add conditional and descriptive statistics wave ([6088ff1](https://www.github.com/danfma/my-sheet/commit/6088ff174ebb1e92f2311bed5675bca646ca1f17))

<a name="2.0.0"></a>
## [2.0.0](https://www.github.com/danfma/my-sheet/releases/tag/v2.0.0) (2026-07-02)

### Breaking Changes

* reorganize AST nodes into semantic namespaces ([50e0486](https://www.github.com/danfma/my-sheet/commit/50e0486c1bcc883fa922be52046abc6bcc9e3d31))

<a name="1.3.0"></a>
## [1.3.0](https://www.github.com/danfma/my-sheet/releases/tag/v1.3.0) (2026-07-02)

### Features

* **functions:** add scalar lookup and reference wave ([e0bd9b1](https://www.github.com/danfma/my-sheet/commit/e0bd9b1417350b54e25f32896e174b3a9cd867c0))

### Bug Fixes

* **lookup:** VLOOKUP returns #VALUE! for col_index_num below 1 ([0fb6af5](https://www.github.com/danfma/my-sheet/commit/0fb6af5388f66925a81148ad7c05b31a034fe3e8))

<a name="1.2.0"></a>
## [1.2.0](https://www.github.com/danfma/my-sheet/releases/tag/v1.2.0) (2026-07-02)

### Features

* **functions:** add information family (NA, IS*, N, T, TYPE, ERROR.TYPE, SHEETS) ([ae49674](https://www.github.com/danfma/my-sheet/commit/ae496747d6654bfef5c73068ca5a4c9be4a48690))
* **functions:** add text formatting, TEXTBEFORE/TEXTAFTER and regex families ([93cd482](https://www.github.com/danfma/my-sheet/commit/93cd48291cdc7a3d5b231d9f46c8851fe29d6086))
* **functions:** add text manipulation family (RIGHT, FIND, SEARCH, ...) ([5b76b04](https://www.github.com/danfma/my-sheet/commit/5b76b04a6665ab2b3254afc24439d8b787075b06))
* **functions:** add TRUE, FALSE, XOR, IFS and SWITCH ([925cf43](https://www.github.com/danfma/my-sheet/commit/925cf437f63638bf6a1b89553795be0a6bf9e1fa))

<a name="1.1.0"></a>
## [1.1.0](https://www.github.com/danfma/my-sheet/releases/tag/v1.1.0) (2026-07-02)

### Features

* **excel:** expande shared formulas no reader (escravas viram formulas reais) ([bbbe60e](https://www.github.com/danfma/my-sheet/commit/bbbe60e0adf17ec4c8f7449e90ceb51bcda38bb0))
* **functions:** add scalar math and trigonometry wave ([1b7fc77](https://www.github.com/danfma/my-sheet/commit/1b7fc77bb86222ec744e9653468ee48153372e03))

<a name="1.0.0"></a>
## [1.0.0](https://www.github.com/danfma/my-sheet/releases/tag/v1.0.0) (2026-07-01)

### Features

* **core:** cache de celula passa a armazenar ComputedValue (Fase 5a) ([d7e0628](https://www.github.com/danfma/my-sheet/commit/d7e0628916c9848658e0c90e42ffb0fff6b19da7))
* **core:** CustomFunction retorna ComputedValue; remove ComputedValue.From ([8443dff](https://www.github.com/danfma/my-sheet/commit/8443dffdff47a1ca2ce9c0e836945a588ef86891))
* **excel:** MergeIntoExcel — injeta valores computados em .xlsx existente ([5e4c42b](https://www.github.com/danfma/my-sheet/commit/5e4c42beb22e2bc9cf8e43119dcdc80f7dd5c84e))
* **excel:** nova lib Danfma.MySheet.Excel com reader ExcelFile.Load (.xlsx -> Workbook) ([c038257](https://www.github.com/danfma/my-sheet/commit/c0382578736f93c223e4277e4b49bf59d7c50dc3))
* **excel:** SaveAsExcel — exporta Workbook para .xlsx (ValuesOnly | Formulas) ([5de9fde](https://www.github.com/danfma/my-sheet/commit/5de9fdec13990db59ff034596fa8806f15d72cc0))
* **expressions:** adiciona ComputedValue e Error (tipos core, aditivo) ([903aadf](https://www.github.com/danfma/my-sheet/commit/903aadfe07f58a4cc8a93b002f630d746147eac0))
* **expressions:** contrato Evaluate + coercao nativa + nos-valor (Fase 2) ([53b71fb](https://www.github.com/danfma/my-sheet/commit/53b71fb66d1c99b336d0a11b2238d8531f6cb853))
* **expressions:** migra agregacao/variadicos/condicionais (Fase 3e) ([256a506](https://www.github.com/danfma/my-sheet/commit/256a506173f5bcc2f9baefd6ee41db6eb1431f6b))
* **expressions:** migra financeiras para Evaluate nativo (Fase 4a) ([044a5e2](https://www.github.com/danfma/my-sheet/commit/044a5e294b9466e12a3498ac31b6c50a73845f47))
* **expressions:** migra lookup/LET/FunctionCall para Evaluate nativo (Fase 4b) ([8d4bb70](https://www.github.com/danfma/my-sheet/commit/8d4bb702fb6aa9c0f2f5d6ae3d4bf009e747fa8a))
* **expressions:** migra math/info escalares para Evaluate nativo (Fase 3c) ([25907f1](https://www.github.com/danfma/my-sheet/commit/25907f157a024eb0343300e853d3ad537902920f))
* **expressions:** migra nos logicos para Evaluate nativo (Fase 3a) ([fa7466f](https://www.github.com/danfma/my-sheet/commit/fa7466f655b7b275cf82ab8089c9f56fb2347d8d))
* **expressions:** migra operadores para Evaluate nativo (Fase 3b) ([6f5457c](https://www.github.com/danfma/my-sheet/commit/6f5457c3e41e8138d1e5bc4d31715d8c878575fe))
* **expressions:** migra texto escalar para Evaluate nativo (Fase 3d) ([c233019](https://www.github.com/danfma/my-sheet/commit/c233019ca7608fa6a3aeda0cd860a4f0735b817b))
* **expressions:** remove Compute; Evaluate:ComputedValue e a unica API ([19b389b](https://www.github.com/danfma/my-sheet/commit/19b389b917613897a97e247ac446f4ecaa965c64))
* **parsing:** FormulaWriter — un-parse de Expression para texto de formula Excel ([6e4b381](https://www.github.com/danfma/my-sheet/commit/6e4b3818d1e331ebf6b3f07a696bfc0129cb34d8))

### Breaking Changes

* **core:** CustomFunction retorna ComputedValue; remove ComputedValue.From ([8443dff](https://www.github.com/danfma/my-sheet/commit/8443dffdff47a1ca2ce9c0e836945a588ef86891))
* **expressions:** remove Compute; Evaluate:ComputedValue e a unica API ([19b389b](https://www.github.com/danfma/my-sheet/commit/19b389b917613897a97e247ac446f4ecaa965c64))

<a name="0.2.0"></a>
## [0.2.0](https://www.github.com/danfma/my-sheet/releases/tag/v0.2.0) (2026-06-29)

### Features

* **financial:** adiciona PMT, PV, FV, NPER, IPMT, PPMT, NPV, RATE e IRR ([da1b9b8](https://www.github.com/danfma/my-sheet/commit/da1b9b865fbe5dd91287d936c71db47c0b500d67))

<a name="0.1.1"></a>
## [0.1.1](https://www.github.com/danfma/my-sheet/releases/tag/v0.1.1) (2026-06-29)

### Bug Fixes

* **lookup:** approximate-match compara chaves de texto em VLOOKUP/MATCH/XLOOKUP ([3460bb3](https://www.github.com/danfma/my-sheet/commit/3460bb3eaae706b5a854ce3fd1b330d82c8ca4aa))

<a name="0.1.0"></a>
## [0.1.0](https://www.github.com/danfma/my-sheet/releases/tag/v0.1.0) (2026-06-25)

### Features

* add conditional/logical functions and text equality comparators ([0468467](https://www.github.com/danfma/my-sheet/commit/04684674478d671fe1948439c7bb73f5f3fd5e8f))
* add custom-function extension mechanism and 23 Excel functions ([e425b79](https://www.github.com/danfma/my-sheet/commit/e425b79136f0084aab51b6b83e6a817af6f036f2))
* add INDEX, MATCH, ROW, ROWS with 2D range access ([863581d](https://www.github.com/danfma/my-sheet/commit/863581dd07850d128a728b7dd871a5aeb1179787))
* add LET and TEXT functions with name bindings ([f9e2964](https://www.github.com/danfma/my-sheet/commit/f9e2964356407aef50628ea68f1b22fd79abe408))
* add RunWithLargeStack for deep dependency chains ([7b75c01](https://www.github.com/danfma/my-sheet/commit/7b75c012bad893f7ac6ba8f7f9af233911a511ba))
* add SHEET function and route lookups through the cache ([7dbf7b7](https://www.github.com/danfma/my-sheet/commit/7dbf7b7b1084f4665fb677f030cc851f2bb2b20d))
* add the & text-concatenation operator ([814d3f3](https://www.github.com/danfma/my-sheet/commit/814d3f37c1047ffab2ed2903c8aaef92cb1859e1))
* add the % (percent) postfix operator ([886186c](https://www.github.com/danfma/my-sheet/commit/886186c564904acdcee477942239dbac95df8d46))
* add the reference-union operator (A1:A3, C1:C3) ([2a19ea7](https://www.github.com/danfma/my-sheet/commit/2a19ea7e3213a04f61e5c7ae1805d3c910dc04f8))
* add VLOOKUP, XLOOKUP, OFFSET and context-aware ROW() ([3cfd752](https://www.github.com/danfma/my-sheet/commit/3cfd752b8cab773e91908baab7f948e5709d799b))
* add Workbook.Save/Load with async overloads ([dfb3b5b](https://www.github.com/danfma/my-sheet/commit/dfb3b5b879e18bf5470d240fcc5f33c33fd15988))
* cache range cells and detect circular references ([ffa27e9](https://www.github.com/danfma/my-sheet/commit/ffa27e9616872320d4c23c15e97abbbda7bc5667))
* Excel cross-type comparison ordering ([6477116](https://www.github.com/danfma/my-sheet/commit/6477116897b1389e022179389aab86dfe317c95c))
* implement core expression functions and parsing logic ([6be60a2](https://www.github.com/danfma/my-sheet/commit/6be60a23688fc42234b10536960d3542ff06a46d))
* initial commit ([d7b930c](https://www.github.com/danfma/my-sheet/commit/d7b930ce84bce95595d49afed6fda9d1224bea99))
* memoize cell values with explicit invalidation ([fea015b](https://www.github.com/danfma/my-sheet/commit/fea015b7ade0a076edb3e10a9605fdd0681de55a))
* range-returning OFFSET, XLOOKUP modes, and omitted arguments ([e4ca37f](https://www.github.com/danfma/my-sheet/commit/e4ca37f1db68b9c20f52f3bd2694bfcd58985459))
* refactor numeric aggregation functions to use folding structures ([8df7d8f](https://www.github.com/danfma/my-sheet/commit/8df7d8f19a063f653b27dde95716665ae117153f))
* support sheet-qualified and absolute ($) cell references ([db36f30](https://www.github.com/danfma/my-sheet/commit/db36f3067e8aef4881ecb75b1f1dac005e25799a))
* TEXT date formats and case-insensitive sheet names ([ba4aa63](https://www.github.com/danfma/my-sheet/commit/ba4aa63ed64d91de0fb93877d95d9dd98496911a))
* update performance benchmarks and optimize numeric aggregation ([87a2523](https://www.github.com/danfma/my-sheet/commit/87a25237e0ad3afc773ce4819294e9c7aff32a6f))

