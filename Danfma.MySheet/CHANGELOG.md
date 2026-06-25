# Change Log

All notable changes to this project will be documented in this file. See [versionize](https://github.com/versionize/versionize) for commit guidelines.

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

