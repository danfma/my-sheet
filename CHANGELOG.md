# Change Log

All notable changes to this project will be documented in this file. See [versionize](https://github.com/versionize/versionize) for commit guidelines.

<a name="3.22.0"></a>
## [3.22.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.22.0) (2026-09-16)

### Features

* **eval:** evaluate array constants as computed arrays ([4b46c4c](https://www.github.com/danfma/my-sheet/commit/4b46c4c680740f0b375e1c41aa01bbda4898d26a))
* **lookup:** index area_num selects an area of a union reference ([ef0d677](https://www.github.com/danfma/my-sheet/commit/ef0d677f2cbaa4e9f24b0f21a29cbe7a00cff4af))
* **parsing:** array constants ([8e7572f](https://www.github.com/danfma/my-sheet/commit/8e7572fe74760389e853dace3a09e3ad15649c33))

### Bug Fixes

* **dates:** workday never throws on out-of-range day counts ([c8ffcfa](https://www.github.com/danfma/my-sheet/commit/c8ffcfabe88ce16d420ee5e84eeb025667e0d3fd))
* **eval:** normalize computed negative zero ([2574408](https://www.github.com/danfma/my-sheet/commit/2574408f6b75346b93e8c3d7082108d3bde006df))
* **lookup:** align empty keys and vectors ([fe2c7bc](https://www.github.com/danfma/my-sheet/commit/fe2c7bc02f2c9a5220dee87b22a9c823e876267b))
* **lookup:** classify resolved absent keys ([e9ddbda](https://www.github.com/danfma/my-sheet/commit/e9ddbda9d1d009b29109d36b8bccbb862f5f0667))
* **lookup:** distinguish absent and empty keys ([c31b0cf](https://www.github.com/danfma/my-sheet/commit/c31b0cf83bee940b4be0746dede7b5736c7fc134))
* **lookup:** match exact empty text keys ([7e1262d](https://www.github.com/danfma/my-sheet/commit/7e1262d88ea451fdf13e5089df67a2a9986d750c))
* **lookup:** match non-text wildcard keys exactly ([a06bfce](https://www.github.com/danfma/my-sheet/commit/a06bfcebed2ef99d963316fa71a01f0511ea656d))
* **lookup:** match reverse absent wildcard order ([20c1f52](https://www.github.com/danfma/my-sheet/commit/20c1f52ae9a7bf63c5aec2e5ac60757971a8409b))
* **lookup:** match wildcards over computed arrays ([0b3c0bd](https://www.github.com/danfma/my-sheet/commit/0b3c0bd157e4d3e1b483972a5b2b566b614864f5))
* **lookup:** preserve computed vector evaluation ([9f252ae](https://www.github.com/danfma/my-sheet/commit/9f252ae4dc4a7ca6c266c60790023fe20b194eaa))
* **lookup:** preserve derived blank key semantics ([f5fde75](https://www.github.com/danfma/my-sheet/commit/f5fde75b0975efcdaf88e8f3509b0fc34800244a))
* **lookup:** preserve selector and structural errors ([aa683fb](https://www.github.com/danfma/my-sheet/commit/aa683fb583eb10be0bbf5e00453751fbeec68659))
* **lookup:** preserve structural vector errors ([5092106](https://www.github.com/danfma/my-sheet/commit/509210624264dd5657edd5d6693a04df462ab8a6))
* **lookup:** preserve wildcard blank key kinds ([a206fd3](https://www.github.com/danfma/my-sheet/commit/a206fd3d16f6bcdcbff0c08d626b3c5a5980f108))
* **lookup:** preserve XMATCH reference arrays ([22b4721](https://www.github.com/danfma/my-sheet/commit/22b47212ae9e41bef758bcd6018f498f90955565))
* **lookup:** resolve wildcard strategy once ([450694e](https://www.github.com/danfma/my-sheet/commit/450694eb6d34b3b34d865e93de686eae9373b3af))
* **lookup:** return ref errors for scalar positions ([efcfae3](https://www.github.com/danfma/my-sheet/commit/efcfae3ed1b98f8cdfe214de9adcc267a9372a94))
* **lookup:** row and column over a computed array ([df876ca](https://www.github.com/danfma/my-sheet/commit/df876ca10d213cbcbd7bbaabcd0f9a972038d963))
* **lookup:** scan absent wildcards symmetrically ([46a47a9](https://www.github.com/danfma/my-sheet/commit/46a47a98d1f93a2c49f350675e75b2faf2d33ad2))
* **lookup:** stream computed XMATCH arrays ([e5730b6](https://www.github.com/danfma/my-sheet/commit/e5730b63d18ac09784fd867470af0a87f174114f))
* **lookup:** unify table lookup semantics ([e9f5171](https://www.github.com/danfma/my-sheet/commit/e9f517157bd2a4395162c6b49f86acfe353e2e1c))
* **lookup:** validate every XMATCH array shape ([d453421](https://www.github.com/danfma/my-sheet/commit/d4534217a56934ca74e5880f2f08053829195a4e))
* **lookup:** validate only the selected index area for a missing sheet ([3301afc](https://www.github.com/danfma/my-sheet/commit/3301afc5e86d5eb0a2fcb73f40da42ab9296778f))
* **parsing:** reject negative zero in array constants ([c724f5d](https://www.github.com/danfma/my-sheet/commit/c724f5d906105e6103ded965b944985e1c7d0138))
* **parsing:** support structured column spans ([e6c6796](https://www.github.com/danfma/my-sheet/commit/e6c6796836844383d4c0425daf70c541e0e8bfd7))
* **text:** preserve unsupported format divergences ([985c46b](https://www.github.com/danfma/my-sheet/commit/985c46bd133d907d2a66df978d7590f8e46d569a))
* **text:** preserve zero-format section semantics ([b72c335](https://www.github.com/danfma/my-sheet/commit/b72c335a45b8a67276cd1120463f8ed03a367949))
* **text:** render selected empty sections ([462bcc9](https://www.github.com/danfma/my-sheet/commit/462bcc97e26ec80346222ed2789af5d7e412f093))

<a name="3.21.0"></a>
## [3.21.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.21.0) (2026-09-15)

### Features

* **excel:** read <table> parts into Workbook.Tables ([15a995f](https://www.github.com/danfma/my-sheet/commit/15a995fbcc9d2e1680df8e156fe7a4270402f4d3))
* **parser:** a structured reference stays on the shared-formula fast path ([e8fa093](https://www.github.com/danfma/my-sheet/commit/e8fa0930d5459223bd5e855b03f8cb07597d78ae))
* **parser:** a structured table reference parses into its node ([6014886](https://www.github.com/danfma/my-sheet/commit/6014886a93ed8f4cdeb5aae9b56ec25af2eefcb2))
* **parser:** bracket token, structured-reference error kinds and the balanced-bracket scanner ([d2fd4e6](https://www.github.com/danfma/my-sheet/commit/d2fd4e6184543161f540259071e92fb36d3afc92))
* **parser:** the lexer reads a bracketed specifier as one raw token ([6a21cc2](https://www.github.com/danfma/my-sheet/commit/6a21cc282246f2a7929061bf440580eb3be87862))
* **parser:** the structured-reference grammar and its canonical writer ([4c39ae6](https://www.github.com/danfma/my-sheet/commit/4c39ae605668376840a5b441624e07272aa685e7))
* **tables:** a bare table name resolves through the name path ([2defd5d](https://www.github.com/danfma/my-sheet/commit/2defd5dd7ea5682b8eb53b40ad52a4d4dfded8ce))
* **tables:** a structured reference streams element by element in the mini-CSE ([4967ba5](https://www.github.com/danfma/my-sheet/commit/4967ba50899256f724b760ce487b4ed5c2f81abd))
* **tables:** ISFORMULA and FORMULATEXT resolve a structured reference ([797025c](https://www.github.com/danfma/my-sheet/commit/797025c5849251cb74f77b7aafd4c90a603da145))
* **tables:** the missing-sheet guard and the static range dep of a structured reference ([d06c43d](https://www.github.com/danfma/my-sheet/commit/d06c43d3e5c7310acd715cd36dafa9aa09d546a3))
* **tables:** the six-area geometry a structured reference resolves through ([c8048a1](https://www.github.com/danfma/my-sheet/commit/c8048a10f773d64ace7e892c210fdbc71331257e))
* **tables:** the TableReference node, TableArea and MemoryPack union tag 327 ([fee97be](https://www.github.com/danfma/my-sheet/commit/fee97be8b976023057d3adc30f2778719f0231b5))

### Bug Fixes

* **eval:** a bare-reference branch of a scalar-condition selector carries its reference ([a8ef927](https://www.github.com/danfma/my-sheet/commit/a8ef927908de76e2f1e51b10df52dfc8d16ba61b))
* **eval:** a lookup value that is a single cell holding an error leads MATCH, XMATCH and XLOOKUP ([5975c63](https://www.github.com/danfma/my-sheet/commit/5975c6351ec900cc83cda94b8248cd8dc625fb03))
* **eval:** a selection over a zero-row source keeps its zero rows, and SUMPRODUCT compares a zero-row shape ([e998533](https://www.github.com/danfma/my-sheet/commit/e998533f9ec26c3af522ab0d0c1ddc9962dca1da))
* **eval:** a single-cell bare-reference IF/CHOOSE branch hands scalar consumers the cell's value ([77aea32](https://www.github.com/danfma/my-sheet/commit/77aea323818ae14abe87669f92938a1a9c57d117))
* **eval:** a structured reference over a zero-row band is an empty reference, not #REF! ([0acb66a](https://www.github.com/danfma/my-sheet/commit/0acb66acff56f1b7f5eae8fc0097b18fb5554a96))
* **eval:** a volatile IF/CHOOSE condition bound to a LET name is drawn exactly once ([f3c09a1](https://www.github.com/danfma/my-sheet/commit/f3c09a15313c7c4173d51f0b5950e58f4f183c86))
* **eval:** admit IF reference selectors in criteria slots ([a3edc84](https://www.github.com/danfma/my-sheet/commit/a3edc8490f8b60779612126df04346aa5a970f1f))
* **eval:** an error-valued argument propagates from the criteria family's range slot and from the resolving consumers ([a542f56](https://www.github.com/danfma/my-sheet/commit/a542f56cdcd991059ae955cf3bd1b074d0937990))
* **eval:** apply negative-index validation to single-cell INDEX tables ([0edca24](https://www.github.com/danfma/my-sheet/commit/0edca246b7a0d60b180842ece362522c0a227248))
* **eval:** bound OFFSET targets to the worksheet grid ([3239dda](https://www.github.com/danfma/my-sheet/commit/3239dda3f496276dc695a8591a442b2d11041a6f))
* **eval:** guard zero-axis INDEX on missing sheets ([a0372df](https://www.github.com/danfma/my-sheet/commit/a0372df892f7e298b42d703eeb2b0eac8027f9f0))
* **eval:** INDEX with a zero row or column returns a reference ([d522027](https://www.github.com/danfma/my-sheet/commit/d52202725a0476a09448d606dab2b577370a78b3))
* **eval:** inherit omitted OFFSET dimensions ([606d594](https://www.github.com/danfma/my-sheet/commit/606d594a46b5a0098a5b25226a670e86aacbb84f))
* **eval:** intersect references in scalar consumers ([e304226](https://www.github.com/danfma/my-sheet/commit/e30422651e6c0051061baa903b083b98e8a30455))
* **eval:** intersect scalar information references ([97ec47d](https://www.github.com/danfma/my-sheet/commit/97ec47d61d39ea42428798b04c970d520146eec0))
* **eval:** keep second-use admission for criteria range snapshots ([9d7419d](https://www.github.com/danfma/my-sheet/commit/9d7419d432f8ff4283421cfdd28a45e26f7e7c9d))
* **eval:** keep structural #REF! for LET-bound missing-sheet criteria ranges ([b57af95](https://www.github.com/danfma/my-sheet/commit/b57af955241cf9509c438c4ea10d64839b427bd9))
* **eval:** lift INDEX and OFFSET references under operators ([e04ae14](https://www.github.com/danfma/my-sheet/commit/e04ae147e65e266ac31e628a12809d9af5756319))
* **eval:** MATCH's approximate path leads on any lookup-value error again, and a 1x1 range holding one counts as a single cell ([4ecd5b5](https://www.github.com/danfma/my-sheet/commit/4ecd5b5a4a8ecc0f4fe0b61d9f7ba37ee7f4c93c))
* **eval:** open-range positional functions address absolute coordinates ([81f228e](https://www.github.com/danfma/my-sheet/commit/81f228ea3ff68acf7c446153a0a1d39b6a7dbcf5))
* **eval:** preserve LET-bound references ([b35ec47](https://www.github.com/danfma/my-sheet/commit/b35ec47f7ad9d0c49fd55de56b9bdbc0626bf51e))
* **eval:** preserve open-range lookup coordinates ([a21ff92](https://www.github.com/danfma/my-sheet/commit/a21ff924bbcd1d0520b76f631197c6e62fc859ff))
* **eval:** refuse a LET-bound xlookup as a criteria reference route ([37b5335](https://www.github.com/danfma/my-sheet/commit/37b5335710f669d703a990142c1aa3d2b9327e3b))
* **eval:** reject selected missing-sheet value ranges ([380c8e0](https://www.github.com/danfma/my-sheet/commit/380c8e02f9bfa1eff1a205c747ceebc81489c33a))
* **eval:** resize SUMIF value ranges ([bbb4e10](https://www.github.com/danfma/my-sheet/commit/bbb4e10279bd664164a217b17582eb6b16b5ea39))
* **eval:** resolve criteria and value selectors through one route classifier ([3f7f8ef](https://www.github.com/danfma/my-sheet/commit/3f7f8ef2b83bf93422e1fa02ae22735b323f5c53))
* **eval:** resolve IF reference selectors in SUMIF sum ranges ([3aacef2](https://www.github.com/danfma/my-sheet/commit/3aacef2933925fe29a668f5fd0bb1e9af11e2aca))
* **eval:** return VALUE for negative INDEX axes ([42c5527](https://www.github.com/danfma/my-sheet/commit/42c5527ef409643b4bd59ac614cac1cea75d68b2))
* **eval:** reuse failed OFFSET resolution values ([7169fb2](https://www.github.com/danfma/my-sheet/commit/7169fb21018b3806e0345a1cabb3782f142c461b))
* **eval:** support negative OFFSET dimensions ([064df31](https://www.github.com/danfma/my-sheet/commit/064df31e463792a45c7c0dc761ed1d793f5292d7))
* **eval:** treat LET chains ending in a reference as structural selector routes ([71abb67](https://www.github.com/danfma/my-sheet/commit/71abb67aa47bd0720c8fc82afb6745d558e30cf5))
* **eval:** validate the final selected reference behind nested IF and LET ([b661979](https://www.github.com/danfma/my-sheet/commit/b661979ccd10b033d6735a6dad549b16fd954606))
* **lookup:** accept one-cell xlookup axes ([127e1f2](https://www.github.com/danfma/my-sheet/commit/127e1f2a7316f9dc6645fec1e9df1e883e1b2875))
* **lookup:** align resolving array semantics ([bfd8702](https://www.github.com/danfma/my-sheet/commit/bfd8702188175cd961a94ef5cad7ec960b46de9a))
* **lookup:** bind xlookup arrays once ([4ebb475](https://www.github.com/danfma/my-sheet/commit/4ebb47511d24478d5c1a3d3a8f0ff0239ae73fa9))
* **lookup:** bind xlookup name arrays ([0cc2bd5](https://www.github.com/danfma/my-sheet/commit/0cc2bd5fb9d26d1cacd008665e12e7592ecce702))
* **lookup:** dereference single-cell lookup values ([6c89daf](https://www.github.com/danfma/my-sheet/commit/6c89dafd61550ff82e09276ab8ad9c3adbdbb38a))
* **lookup:** keep computed xlookup rows and columns as arrays ([25da763](https://www.github.com/danfma/my-sheet/commit/25da763c846ecd9c1efd48ed82713bf67296d66d))
* **lookup:** map xlookup results by search axis ([e457041](https://www.github.com/danfma/my-sheet/commit/e4570413cc3cf2eb8a823b58c033f4f518fdbdf3))
* **lookup:** memoize the xlookup selection once per evaluation ([3dd0dcf](https://www.github.com/danfma/my-sheet/commit/3dd0dcfa4751d1edff106aba171f68a890122b60))
* **lookup:** preserve direct error table precedence ([c080de9](https://www.github.com/danfma/my-sheet/commit/c080de94c47c36cdca81642543478e243ca738a2))
* **lookup:** propagate lookup-value errors over single-cell tables ([f65f665](https://www.github.com/danfma/my-sheet/commit/f65f66570f06a00699d2db8cbef8897930401386))
* **lookup:** report a non-reference xlookup result in reference-only slots ([979effc](https://www.github.com/danfma/my-sheet/commit/979effc99032db5cf93c550737594e82ef90216d))
* **lookup:** resolve reference-returning xlookup at shared reading paths ([07b43d0](https://www.github.com/danfma/my-sheet/commit/07b43d023d9c857177704c83ffbb18eacfa36d8e))
* **lookup:** resolve the match lookup array once ([489b592](https://www.github.com/danfma/my-sheet/commit/489b5924aee7a20348b66260b2f4cb6d7ae981c7))
* **lookup:** select open-range results by cell coordinate ([5e4848d](https://www.github.com/danfma/my-sheet/commit/5e4848d13253818c6337d3cb441f0c7310a4800a))
* **lookup:** treat a one-cell xlookup lookup array as a row ([f76626c](https://www.github.com/danfma/my-sheet/commit/f76626c2da71a5a3e3840bbf8532b1a9f8b9e2ac))
* **lookup:** validate single-cell tables ([cfbdfbc](https://www.github.com/danfma/my-sheet/commit/cfbdfbc442890b0073ed5d0986e8fa17d012becd))
* **parser:** an apostrophe escapes only the five specials, it is literal before anything else ([8f94b18](https://www.github.com/danfma/my-sheet/commit/8f94b1838edabb95e6555f04553cb7c8bd0e3f01))
* **parser:** defined names use Excel's grid-bounded cell-reference rule ([3e89cfb](https://www.github.com/danfma/my-sheet/commit/3e89cfbc287c1d07e75f79d65f825db09c8111bf))
* **parser:** the nameless-bracket message stops asserting which shape it is ([f739397](https://www.github.com/danfma/my-sheet/commit/f739397d4f874b6e35ba5fe7f6a7246fcb5b7b4a))
* **parsing:** bound deleted reference continuations ([82b9de0](https://www.github.com/danfma/my-sheet/commit/82b9de047029700feea7d0344ba2d4f660e27764))
* **parsing:** classify deleted reference continuations ([bc753f9](https://www.github.com/danfma/my-sheet/commit/bc753f9aeb77f3af46eded1d39558c833a849b23))
* **parsing:** read error literals in formula text ([5c7eeab](https://www.github.com/danfma/my-sheet/commit/5c7eeab587ced11945022bc74f49fef95747ee49))

<a name="3.20.0"></a>
## [3.20.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.20.0) (2026-09-11)

### Features

* **arrays:** a defined name, CHOOSE and unary + carry a computed array too ([710f3c3](https://www.github.com/danfma/my-sheet/commit/710f3c35e91615e214fc256fdf4c99651c950ce4))
* **arrays:** a LET binding carries a computed array, built once and read by every consumer ([8daaa56](https://www.github.com/danfma/my-sheet/commit/8daaa5639d22ff6aa4e9b6e78a26bfb81d8c257d))
* **arrays:** COUNTA, CONCAT and TEXTJOIN stream a computed array element by element ([7e57071](https://www.github.com/danfma/my-sheet/commit/7e57071707922e1218cf5c6e2de20bbc7e968c04))
* **arrays:** FILTER, SORT and UNIQUE as mini-CSE producers over one axis-selection operand ([41272eb](https://www.github.com/danfma/my-sheet/commit/41272eb31d5cd9e987c5dc14b277508269beff94))
* **arrays:** register FILTER, SORT, UNIQUE and SEQUENCE as Consumes entries with union tags 323-326 ([96bcf3a](https://www.github.com/danfma/my-sheet/commit/96bcf3a1f4b27aed803cd1b35289b549d5fb3f69))
* **arrays:** ROWS and COLUMNS answer the shape of a computed array through the mini-CSE gate ([ce61bce](https://www.github.com/danfma/my-sheet/commit/ce61bce1eb83c3aa52a3f6dd683b71c009a4d80c))
* **arrays:** SEQUENCE as the first mini-CSE producer, with its operand under the shape invariant ([dc3fa4a](https://www.github.com/danfma/my-sheet/commit/dc3fa4a387bb5beb1e74c8a7256a65f48270ac3a))
* **arrays:** the producer contract — IArrayProducer, its two arms, FirstElement and ArrayShaping ([c76acfb](https://www.github.com/danfma/my-sheet/commit/c76acfb2fb297756a76359b4f56fbccdde0051d6))
* **errors:** #CALC! as the eighth error code, and ERROR.TYPE answers 14 for it ([380e046](https://www.github.com/danfma/my-sheet/commit/380e04659e5412f1effef1b7ed90b581b9dd6830))

### Bug Fixes

* **arrays:** a volatile condition could still collapse a producer under a scalar-condition IF ([1658a0a](https://www.github.com/danfma/my-sheet/commit/1658a0aef45cdae71a564bea998e54bbd025030d))
* **arrays:** probe an IF branch by the CONDITION's kind, restoring the probe/build lockstep ([46d1a85](https://www.github.com/danfma/my-sheet/commit/46d1a855247a012228f7ec67de0e9f300a879f13))
* **arrays:** ReferenceGuard sees a missing sheet through FILTER, SORT and UNIQUE ([cfb72cc](https://www.github.com/danfma/my-sheet/commit/cfb72cc2e71aee1bccb6d9976661e11983dac533))
* **arrays:** stream a producer under a SCALAR-condition IF instead of collapsing it ([111dd2c](https://www.github.com/danfma/my-sheet/commit/111dd2c4d07bfa9cbb29c02bd71b257455e12bca))
* **eval:** accept the text TRUE/FALSE in IF, NOT and IFS' condition slots ([90df14b](https://www.github.com/danfma/my-sheet/commit/90df14b4603c110757374e00f8c98e0276191516))

<a name="3.19.0"></a>
## [3.19.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.19.0) (2026-09-10)

### Features

* **tables:** register tables on the Workbook as its third serialized member ([0887770](https://www.github.com/danfma/my-sheet/commit/0887770f69dc223d8de9a24f56c42ad68381d993))
* **tables:** Table record with derived geometry and a column index memoized off the record ([41adb54](https://www.github.com/danfma/my-sheet/commit/41adb54e6a4a939182e2289d9ce460cdadfa2217))
* **tables:** table-name rule, Table.Validate and the grid-bounded cell-reference check ([873142b](https://www.github.com/danfma/my-sheet/commit/873142bfe4ca615ba6ff9ad8686b52cf6b93a1e1))

### Bug Fixes

* **parsing:** bound TryParseColumn's accumulator and add CellAddress.TryParseA1 ([7142472](https://www.github.com/danfma/my-sheet/commit/7142472da1628d19f91f8de6f963bbb16fb05216))
* **recalc:** tratar mudança de definição como invalidação total ([40646c5](https://www.github.com/danfma/my-sheet/commit/40646c503eead76a3161c3a69ab2a4b50a433821))

<a name="3.18.0"></a>
## [3.18.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.18.0) (2026-09-10)

### Features

* **criteria:** a computed array in a criteria-family range slot is #REF! ([cf8707b](https://www.github.com/danfma/my-sheet/commit/cf8707b852bf5c7f47e590e8fde040caba2951d4))
* **eval:** a defined name in an array position is whatever it is bound to ([8ae501d](https://www.github.com/danfma/my-sheet/commit/8ae501d98cae62deefc8e97a5a839dc130aaf51a))

<a name="3.17.0"></a>
## [3.17.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.17.0) (2026-09-10)

### Features

* **eval:** AGGREGATE(function_num, options, ref1, [k]) ([8087a8b](https://www.github.com/danfma/my-sheet/commit/8087a8be1db9118f45b7c61059fc6478ac41cf72))
* **eval:** broadcast vector operands in the mini-CSE, Excel's per-axis rule ([1882888](https://www.github.com/danfma/my-sheet/commit/18828880ce1402b852598f072cc1b18bcfd1ec39))
* **eval:** composites project through Broadcasting.TryProject and IF folds all three shapes ([2ea3537](https://www.github.com/danfma/my-sheet/commit/2ea3537c432d0591cd2814ea7ab718b166d8f7e7))
* **eval:** implicit intersection at the cell boundary ([b8fa6b6](https://www.github.com/danfma/my-sheet/commit/b8fa6b60b4f2be7c2608c2d401491d8680cfa05e))
* **eval:** KthValueStreaming can skip error elements ([745e6fc](https://www.github.com/danfma/my-sheet/commit/745e6fcf1472a585e205541b6d11af44da376e56))
* **eval:** lift unary '-'/'%' and pure-scalar built-ins over arrays in the mini-CSE ([725dea5](https://www.github.com/danfma/my-sheet/commit/725dea5516d90f744ee351875b784ba561acbcd8))
* **eval:** ROW/COLUMN over a name or a reference inside the mini-CSE ([7fa24d2](https://www.github.com/danfma/my-sheet/commit/7fa24d218af7581d9b856a1a26691cf369aae589))
* **eval:** ROW/COLUMN over any reference-producing expression ([d6df57a](https://www.github.com/danfma/my-sheet/commit/d6df57a82afdce68b7506a56e24f92c24f931fef))
* **parser:** classify registry entries for array lifting ([2c7acc6](https://www.github.com/danfma/my-sheet/commit/2c7acc6c22aceb0c47ead4e75294d19d81b595b9))
* **parser:** mark the 180 pure-scalar built-ins as Elementwise ([b3ff90d](https://www.github.com/danfma/my-sheet/commit/b3ff90d370333bd8905ac353bc33530e2d82591c))

### Bug Fixes

* **eval:** a whole-argument error propagates through AGGREGATE's ignore-errors bit ([b83dd7f](https://www.github.com/danfma/my-sheet/commit/b83dd7fd8615d9939e11fe38eaa23a0fb876e08d))
* **eval:** count days on serials and 30/360 on the Lotus calendar ([b60c3ca](https://www.github.com/danfma/my-sheet/commit/b60c3ca2630c23f729fe9f1ca1c703a3709daf11))
* **eval:** DATEVALUE rejects pre-1900 dates and parses the phantom Feb 29 ([e5ffe07](https://www.github.com/danfma/my-sheet/commit/e5ffe07e673d0ae867694d24cf26963a951f89a1))
* **eval:** discount XNPV and XIRR on serial differences ([d5bf3e8](https://www.github.com/danfma/my-sheet/commit/d5bf3e8f95dd6d27b84c4144493adade8c432ab9))
* **eval:** keep a malformed range id out of the cell boundary ([57926e4](https://www.github.com/danfma/my-sheet/commit/57926e457e41dcd83164c1519fac2036b5d0e959))
* **eval:** place serial 1 on 1900-01-01 in the central date map ([db26d57](https://www.github.com/danfma/my-sheet/commit/db26d5707783892d5a95bd2012f2050c232a98b5))
* **eval:** re-check the resolved sheet in ROWS/COLUMNS/AREAS ([02cae13](https://www.github.com/danfma/my-sheet/commit/02cae1394b57126da10d738304712838c5b8fba0))
* **eval:** ROW/COLUMN of a rectangle are vectors, not MxN rectangles ([1ed60b1](https://www.github.com/danfma/my-sheet/commit/1ed60b1d858facded6d9fc682512c20ead5ce048))
* **eval:** SUBTOTAL and AGGREGATE 1-13 reject a computed-array argument, as Excel does ([34a52ab](https://www.github.com/danfma/my-sheet/commit/34a52ab1cfd97e8eb49c355dde54cf716eedbe62))
* **eval:** SUBTOTAL folds a computed-array argument ([3b3b798](https://www.github.com/danfma/my-sheet/commit/3b3b798f9614e46d80afc9748373a24c892ce1e8))
* **eval:** SUMPRODUCT compares every shape, not just the first argument's ([fae04cc](https://www.github.com/danfma/my-sheet/commit/fae04cc2e16b99e979d8396a73eed91353db83ca))
* **eval:** SUMPRODUCT reads computed arrays and honours the dimension rule ([b9ddfc0](https://www.github.com/danfma/my-sheet/commit/b9ddfc01dea6aef500d4cf1b3af90aae3e00f633))
* **eval:** TEXT prints Excel's own 1900 calendar ([0e4ad86](https://www.github.com/danfma/my-sheet/commit/0e4ad86ff01e6e21b15ef146ff3a727027321d98))
* **eval:** the nested-aggregate skip sees through a shared-formula slave ([afec6e9](https://www.github.com/danfma/my-sheet/commit/afec6e9a027565c7940e377bb49fce9c82ebb8e7))
* **eval:** throw, not #VALUE!, on Apply's unreachable Plus arm ([77c2ca2](https://www.github.com/danfma/my-sheet/commit/77c2ca20bf98cd21e99db82960b48d7c423eb7b7))
* **eval:** walk serials in the working-day family, on the real calendar ([5e6c7bd](https://www.github.com/danfma/my-sheet/commit/5e6c7bd3cbad94e2a09ed4dc7c0a6cd0856a4136))
* **eval:** WEEKDAY walks the Lotus weekday, EDATE/EOMONTH stop at day zero ([fb0dc7c](https://www.github.com/danfma/my-sheet/commit/fb0dc7c156018456f5b8a6c7b31765151ffc1350))
* **eval:** YEAR/MONTH/DAY answer Excel's day zero ([e221a92](https://www.github.com/danfma/my-sheet/commit/e221a9297af4cb96958e93287863f80dee633005))
* **excel:** degrade an unparsable formula or literal per cell instead of aborting the load ([507477b](https://www.github.com/danfma/my-sheet/commit/507477b2395b54f85207ead2a00fa636e5fded75))

<a name="3.16.1"></a>
## [3.16.1](https://www.github.com/danfma/my-sheet/releases/tag/v3.16.1) (2026-09-09)

<a name="3.16.0"></a>
## [3.16.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.16.0) (2026-09-08)

### Features

* **parser:** keep the legacy ParseException(message, position) constructor ([fbd360e](https://www.github.com/danfma/my-sheet/commit/fbd360e6e71bb4427c8e75892df9ba079d8c62b9))
* **parser:** structured ParseException (Kind, Token, Position) ([39817d4](https://www.github.com/danfma/my-sheet/commit/39817d487dc2b80530a3d1056007992484f3e82d))

### Bug Fixes

* **eval:** LET, CHOOSE and defined names capture ranges as reference values ([378ad5c](https://www.github.com/danfma/my-sheet/commit/378ad5cfadd932878632d4b1c00b552fa06f84af))
* **eval:** unary + is a type-preserving no-op ([90771b1](https://www.github.com/danfma/my-sheet/commit/90771b12f4da731faebc3a1f47747b247fa925fb))
* **eval:** unary + passes range nodes through to range consumers ([4c85a62](https://www.github.com/danfma/my-sheet/commit/4c85a6205f36d0938e5ef7f62e5e942860e5f21d))
* **parser:** accept absolute row endpoints ($1:$1) in whole-row ranges ([a20e9bd](https://www.github.com/danfma/my-sheet/commit/a20e9bd76b596c932845bda1aa153b56576256b6))
* **parser:** reject absolute row endpoints that overflow int ([9dccf7e](https://www.github.com/danfma/my-sheet/commit/9dccf7ecdb46061e0c8f9aef04226b297f0c6230))

<a name="3.15.0"></a>
## [3.15.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.15.0) (2026-07-13)

### Features

* **serialization:** MSWM container v3 — fixed-chunk streamed Brotli, pooled decompress ([78d0af4](https://www.github.com/danfma/my-sheet/commit/78d0af4c7c0333eb8ec7cdffa23db8ef0981c587))

<a name="3.14.0"></a>
## [3.14.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.14.0) (2026-07-11)

### Features

* **recalc:** integrate shared-formula delta nodes from spike ([bdf1385](https://www.github.com/danfma/my-sheet/commit/bdf138597ed42ef343db262c515111b21a5b8685))
* **recalc:** shared-formula slaves as delta nodes over a shared master tree ([d31d338](https://www.github.com/danfma/my-sheet/commit/d31d3387934a3ed974198f60c94086d73b293746))

### Bug Fixes

* **dirty-graph:** extract effective dependencies for shared-formula nodes ([246a61d](https://www.github.com/danfma/my-sheet/commit/246a61d94ed334f316c58fd1c222020dd2833971))
* **recalc:** teach the remaining reference-pattern sites about anchored nodes ([e26c5b9](https://www.github.com/danfma/my-sheet/commit/e26c5b984b60d023d7b2ce5cf4c2104fdcf168cc))

<a name="3.13.0"></a>
## [3.13.0](https://www.github.com/danfma/my-sheet/releases/tag/v3.13.0) (2026-07-11)

### Features

* **serialization:** opt-in Pipelines I/O to bound Save/Load allocations ([aa1b6f9](https://www.github.com/danfma/my-sheet/commit/aa1b6f9478258d1ae50d853a1b6d310beb3c2db8))
* **serialization:** streaming LoadAsync, single-pass container, opt-in pipelines I/O ([0bb61d8](https://www.github.com/danfma/my-sheet/commit/0bb61d8896463634b29105a7f8d2f6d1dc6c6511))

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

