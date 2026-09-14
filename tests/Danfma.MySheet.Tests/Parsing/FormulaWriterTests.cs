using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet.Tests.Parsing;

public class FormulaWriterTests
{
    private const string ContextSheet = "Sheet1";

    private static Expression Parse(string formula) =>
        ExpressionParser.Parse("=" + formula, new Sheet { Name = ContextSheet });

    private static byte[] Structure(Expression expression) =>
        MemoryPackSerializer.Serialize(expression);

    // Fórmulas já em forma canônica: o un-parse devolve o MESMO texto (round-trip por string exata).
    [Test]
    [Arguments("1+2*3")]
    [Arguments("(1+2)*3")]
    [Arguments("1-(2-3)")]
    [Arguments("2^3^4")] // '^' é right-associative: parseia como 2^(3^4) e re-imprime sem parênteses
    [Arguments("(2^3)^4")]
    [Arguments("-2^3")] // prefixo liga mais forte que '^' (semântica Excel): (-2)^3
    [Arguments("-(2^3)")]
    [Arguments("-A1")]
    [Arguments("10%")]
    [Arguments("(A1+B1)%")]
    [Arguments("A1%*2")]
    [Arguments("1<=2")]
    [Arguments("A1<>\"\"")]
    [Arguments("A1&\" \"&B1")]
    [Arguments("\"a\"\"b\"&C1")] // aspas internas escapadas como ""
    [Arguments("TRUE")]
    [Arguments("SUM(A1:B10)")]
    [Arguments("SUM(A1:A3,C1,5)")]
    // Fase 8 (lifting elementwise): a árvore não ganhou nó nenhum, então o writer alcança os argumentos da
    // função lifted pelo MESMO acessor do registry (FormulaWriter.Call) e o texto é idêntico ao de entrada.
    [Arguments("SUM(LEN(A1:A3))")]
    [Arguments("SUM(-(A1:A3>1))")]
    [Arguments("SUM((A1:A3,C1:C3))")] // união de ranges
    [Arguments("IF(A1>0,\"yes\",\"no\")")]
    [Arguments("IFERROR(1/0,42)")]
    [Arguments("VLOOKUP(A1,B1:C10,2,FALSE)")]
    [Arguments("INDIRECT(\"A1\")")] // regressão: INDIRECT não tinha arm no FormulaWriter.Call
    [Arguments("INDIRECT(\"A1\",TRUE)")]
    [Arguments("COUNTIF(A1:A10,\">5\")")]
    [Arguments("LET(x,1,y,2,x+y)")]
    [Arguments("SHEET()")]
    [Arguments("MYFN(1,A1)")] // função custom (FunctionCall)
    [Arguments("ROUND(1.234,2)")]
    [Arguments("CEILING.MATH(-5.5,2,-1)")] // nome de função com ponto
    [Arguments("PI()")]
    [Arguments("Sheet2!A1+1")]
    [Arguments("Sheet2!A1:B2")]
    [Arguments("'My Sheet'!A1")]
    // Whole-column / whole-row and one-sided open references.
    [Arguments("SUM(A:A)")]
    [Arguments("SUM(A:C)")]
    [Arguments("SUM(1:1)")]
    [Arguments("SUM(1:5)")]
    [Arguments("SUM(A2:A)")]
    [Arguments("SUM(A:A10)")]
    [Arguments("SUM(A1:C)")]
    [Arguments("SUM(Sheet2!A:A)")]
    // Fase 4: referências estruturadas (tabela). São as quinze linhas canônicas do item 16 mais a chamada de
    // função, agora entrando pelo PARSER — até T5 elas viviam como árvores montadas à mão logo abaixo, e
    // nenhum valor esperado mudou na mudança. A grafia canônica é medida em Aspose.Cells 26.6.0, entrada
    // PLAIN, 2026-09-11: colchete SIMPLES sempre, `'` antes de cada um de `[ ] # ' @`, e colchete duplo só
    // para nome de coluna com espaço nas pontas (`Tabela1[[ Col ]]`).
    [Arguments("Tabela1[Valor]")]
    [Arguments("Tabela1[#All]")]
    [Arguments("Tabela1[#Data]")]
    [Arguments("Tabela1[#Headers]")]
    [Arguments("Tabela1[#Totals]")]
    [Arguments("Tabela1[[#All],[Valor]]")]
    [Arguments("Tabela1[[#Headers],[#Data]]")]
    [Arguments("Tabela1[[#Data],[#Totals]]")]
    [Arguments("Tabela1[[#Headers],[#Data],[% Comissao]]")]
    [Arguments("Tabela1[[#Data],[#Totals],[Valor]]")]
    [Arguments("Tabela1[Sales Amount]")]
    [Arguments("Tabela1[Total (USD)]")]
    [Arguments("Tabela1['#OfItems]")]
    [Arguments("Tabela1[[ Col ]]")]
    [Arguments("Tabela1[Rev'#1]")]
    [Arguments("SUM(Tabela1[Valor])")]
    // Item 43 (sweep 31-35-43): an error literal round-trips through its own AST node (ErrorValue), which
    // FormulaWriter already had a case for — this is the PARSER side newly reaching it.
    [Arguments("#REF!")]
    [Arguments("#N/A")]
    [Arguments("SUM(#REF!)")]
    [Arguments("MATCH(#REF!,#REF!,0)")]
    [Arguments("A1:#REF!")]
    [Arguments("#REF!+1")]
    public async Task RoundTrips_CanonicalText(string formula)
    {
        await Assert.That(Parse(formula).ToFormula(ContextSheet)).IsEqualTo(formula);
    }

    // Texto não-canônico (espaços, referência à própria planilha) normaliza, mas a ÁRVORE re-parseada
    // é estruturalmente idêntica (comparação por serialização MemoryPack).
    [Test]
    [Arguments("1 + 2 * 3", "1+2*3")]
    [Arguments("SUM( A1 : A2 )", "SUM(A1:A2)")]
    [Arguments("Sheet1!A1+1", "A1+1")] // referência ao próprio contexto fica sem qualificação
    // Fase 4: os sete pares de normalização do item 16, com as três direções que a Ruling 2 inverteu já
    // corrigidas (medido em Aspose.Cells 26.6.0, entrada PLAIN, 2026-09-11). `[[Valor]]` perde o par de
    // colchetes interno, o espaço decorativo depois da vírgula cai, e `[#totals]` é canonizado. As
    // invertidas: `Tabela1[ Col ]` vira `Tabela1[[ Col ]]` porque o oráculo NÃO apara o espaço
    // (`[ Padded ]` volta como `[[ Padded ]]`), `Tabela1[[Sales Amount]]` volta como
    // `Tabela1[Sales Amount]` (e não o contrário), e `Tabela1[Rev#1]` volta em colchete SIMPLES com o
    // `#` escapado. DIVERGÊNCIA registrada, e não medição do oráculo: os pares `[[#Data],[Valor]]` ->
    // `[Valor]` (e o gêmeo com espaço) são normalização DO MySheet — o oráculo PRESERVA o `[#Data]` ao
    // lado de uma coluna (`=SUM(Tabela1[[#Data],[Valor]])` volta idêntico; só o espaço decorativo cai) e
    // canoniza apenas a ORDEM (`[[Valor],[#Data]]` -> `[[#Data],[Valor]]`; review probe oracle-full3,
    // linhas 6-7 e 48). O colapso é inevitável no modelo — o nó `(Valor, Data)` não distingue a origem —
    // e semanticamente inocente, mas FORMULATEXT/export diverge: work item sob P0, não paridade.
    [Arguments("Tabela1[[Valor]]", "Tabela1[Valor]")]
    [Arguments("Tabela1[[#Data],[Valor]]", "Tabela1[Valor]")]
    [Arguments("Tabela1[[#Data], [Valor]]", "Tabela1[Valor]")]
    [Arguments("Tabela1[#totals]", "Tabela1[#Totals]")]
    [Arguments("Tabela1[ Col ]", "Tabela1[[ Col ]]")]
    [Arguments("Tabela1[[Sales Amount]]", "Tabela1[Sales Amount]")]
    [Arguments("Tabela1[Rev#1]", "Tabela1[Rev'#1]")]
    // Item 43: an error literal canonicalizes to upper-case regardless of input case (matches Aspose's
    // own round trip, measured 2026-09-14: `#ref!` -> formula text `#REF!`).
    [Arguments("#ref!", "#REF!")]
    [Arguments("#n/a", "#N/A")]
    // A sheet-qualified error literal drops the qualifier: the value is #REF! either way once the
    // reference is broken, and Aspose's own object model already discards sheet qualifiers on OTHER
    // shapes it cannot represent structurally (Data!Tabela1[Valor] above). Aspose ITSELF keeps
    // `Sheet1!#REF!` verbatim on round trip (measured) — this is a deliberate divergence, documented in
    // docs/workbook-and-expressions.md (the "Formula text written back..." Parsing bullet) and
    // docs/excel-interop.md (Scope and limitations), both twins — not parity, because ErrorValue
    // carries no sheet and the VALUE never depends on the qualifier.
    [Arguments("Sheet1!#REF!", "#REF!")]
    // Round 2 (I-3): #REF! as a deleted-sheet qualifier consumes the reference glued after it and
    // discards it the same way — `#REF!A1` and the redundant-bang spelling `#REF!!A1` (which Aspose
    // itself already normalizes to `#REF!A1` on ITS OWN round trip, measured) both drop to bare `#REF!`
    // here, same divergence, same doc bullets.
    [Arguments("#REF!A1", "#REF!")]
    [Arguments("#REF!!A1", "#REF!")]
    public async Task NormalizesEquivalentText(string input, string canonical)
    {
        var expression = Parse(input);
        var written = expression.ToFormula(ContextSheet);

        await Assert.That(written).IsEqualTo(canonical);
        await Assert.That(Structure(Parse(written)).SequenceEqual(Structure(expression))).IsTrue();
    }

    // === Fase 4: o arm do TableReference (item 12) =======================================================

    // Sem este arm o `default:` do FormulaWriter estoura NotSupportedException, ou seja FORMULATEXT e toda
    // exportação em modo Formulas morrem em qualquer célula com referência estruturada.
    //
    // As vinte e duas linhas do item 16 (quinze canônicas + sete pares de normalização) vivem agora nas duas
    // listas de [Arguments] acima, medidas em Aspose.Cells 26.6.0, entrada PLAIN, 2026-09-11: um probe FORA
    // da suíte confirmou que as 102 grafias que este writer consegue emitir (17 nomes de coluna, o nulo
    // incluído, x 6 áreas, todas distintas) voltam do oráculo byte a byte iguais — o número 102 é dessa
    // medição, não da contagem de nenhum teste. A tabela COMPLETA de grafias do writer, incluindo as que o
    // parser não produz, é de StructuredReferenceSyntaxTests.Write_RendersTheCanonicalSpelling; o que sobra
    // aqui são os slots em que o nó entra e as travessias que RoundTrips_CanonicalText não alcança.

    // O átomo nunca é parentizado: Precedence devolve AtomPrecedence pelo `_ =>`, então o nó entra cru em
    // qualquer slot de operando. `SUM(Tabela1[Valor])` é a linha canônica do item 16 que envolve o nó numa
    // chamada de função.
    [Test]
    public async Task TableReference_IsAnAtom_InEveryOperandSlot()
    {
        Expression reference = new TableReference("Tabela1", "Valor", TableArea.Data);

        await Assert
            .That(new Sum([reference]).ToFormula(ContextSheet))
            .IsEqualTo("SUM(Tabela1[Valor])");
        await Assert
            .That(
                new BinaryOperation(
                    BinaryOperator.Multiply,
                    reference,
                    new NumberValue(2)
                ).ToFormula(ContextSheet)
            )
            .IsEqualTo("Tabela1[Valor]*2");
        await Assert
            .That(new UnaryOperation(UnaryOperator.Negate, reference).ToFormula(ContextSheet))
            .IsEqualTo("-Tabela1[Valor]");
    }

    // Uma referência estruturada não tem componente de posição, então o deltaRow/deltaColumn ambiente de uma
    // fórmula compartilhada é ignorado — é por isso que ela continua no caminho rápido do master
    // compartilhado (AnchoredFormulaSupport).
    [Test]
    public async Task TableReference_IgnoresTheSharedFormulaDelta()
    {
        Expression reference = new TableReference("Tabela1", "Valor", TableArea.Data);
        Expression slave = new SharedFormulaSlave(reference, 1000, 7);

        await Assert.That(slave.ToFormula(ContextSheet)).IsEqualTo("Tabela1[Valor]");
    }

    // O nó atravessa MemoryPack (tag de união 327, a primeira livre depois da Fase 7) e o clone re-renderiza
    // igual, que é a metade da invariante que NormalizesEquivalentText prova para os outros nós.
    [Test]
    [Arguments(null, TableArea.Data)]
    [Arguments(null, TableArea.HeadersAndData)]
    [Arguments("Valor", TableArea.Data)]
    [Arguments("#OfItems", TableArea.All)]
    [Arguments(" Col ", TableArea.DataAndTotals)]
    public async Task TableReference_SurvivesMemoryPack(string? column, TableArea area)
    {
        Expression reference = new TableReference("Tabela1", column, area);

        var clone = MemoryPackSerializer.Deserialize<Expression>(Structure(reference));

        await Assert.That(clone).IsEqualTo(reference);
        await Assert
            .That(clone!.ToFormula(ContextSheet))
            .IsEqualTo(reference.ToFormula(ContextSheet));
    }

    // Uma chamada mínima válida por função built-in do Parser: o texto re-parseado tem que produzir
    // exatamente a mesma árvore (garante que o mapa nó→nome cobre todas as funções).
    [Test]
    public async Task AllBuiltInFunctions_RoundTripStructurally()
    {
        string[] corpus =
        [
            "SUM(1)",
            "AVERAGE(1)",
            "MIN(1)",
            "MAX(1)",
            "COUNT(1)",
            "IF(TRUE,1,2)",
            "AND(TRUE)",
            "OR(TRUE)",
            "NOT(TRUE)",
            "IFERROR(1,2)",
            "IFNA(1,2)",
            "INT(1.5)",
            "ROUND(1.234,2)",
            "ROUNDUP(1.2,0)",
            "ABS(1)",
            "ISNUMBER(1)",
            "ISBLANK(A1)",
            "UPPER(\"a\")",
            "LOWER(\"A\")",
            "TRIM(\" a \")",
            "LEN(\"a\")",
            "LEFT(\"ab\",1)",
            "MID(\"abc\",2,1)",
            "VALUE(\"1\")",
            "TEXT(1,\"0\")",
            "CONCAT(\"a\")",
            "CONCATENATE(\"a\",\"b\")",
            "TEXTJOIN(\",\",TRUE,\"a\")",
            "COUNTA(A1:A2)",
            "COUNTBLANK(A1:A2)",
            "COUNTIF(A1:A2,1)",
            "COUNTIFS(A1:A2,1)",
            "SUMIF(A1:A2,1)",
            "SUMIFS(A1:A2,B1:B2,1)",
            "ROWS(A1:A2)",
            "ROW()",
            "MATCH(1,A1:A3)",
            "INDEX(A1:B2,1,1)",
            "VLOOKUP(1,A1:B2,2)",
            "XLOOKUP(1,A1:A2,B1:B2)",
            "OFFSET(A1,1,1)",
            "LET(x,1,x)",
            "SHEET()",
            "PMT(0.05,10,1000)",
            "PV(0.05,10,100)",
            "FV(0.05,10,100)",
            "NPER(0.05,100,1000)",
            "IPMT(0.05,1,10,1000)",
            "PPMT(0.05,1,10,1000)",
            "NPV(0.1,A1:A3)",
            "RATE(10,100,1000)",
            "IRR(A1:A3)",
            // Onda 1 — Math & Trigonometria escalar.
            "SQRT(4)",
            "POWER(2,3)",
            "EXP(1)",
            "LN(1)",
            "LOG(8,2)",
            "LOG10(100)",
            "SQRTPI(1)",
            "ROUNDDOWN(1.9,0)",
            "TRUNC(1.9)",
            "MROUND(10,3)",
            "CEILING(2.5,1)",
            "CEILING.MATH(6.7)",
            "CEILING.PRECISE(6.7)",
            "ISO.CEILING(6.7)",
            "FLOOR(3.7,2)",
            "FLOOR.MATH(6.7)",
            "FLOOR.PRECISE(6.7)",
            "EVEN(1)",
            "ODD(2)",
            "MOD(3,2)",
            "QUOTIENT(5,2)",
            "SIGN(-5)",
            "PI()",
            "PRODUCT(A1:A2)",
            "SUMSQ(1,2)",
            "MULTINOMIAL(2,3)",
            "SERIESSUM(1,0,1,A1:A2)",
            "FACT(5)",
            "FACTDOUBLE(6)",
            "COMBIN(8,2)",
            "COMBINA(4,3)",
            "GCD(24,36)",
            "LCM(24,36)",
            "SIN(1)",
            "COS(1)",
            "TAN(1)",
            "COT(1)",
            "SEC(1)",
            "CSC(1)",
            "ASIN(1)",
            "ACOS(1)",
            "ATAN(1)",
            "ATAN2(1,1)",
            "ACOT(1)",
            "SINH(1)",
            "COSH(1)",
            "TANH(1)",
            "COTH(1)",
            "SECH(1)",
            "CSCH(1)",
            "ASINH(1)",
            "ACOSH(2)",
            "ATANH(0.5)",
            "ACOTH(6)",
            "DEGREES(PI())",
            "RADIANS(180)",
            "BASE(7,2)",
            "DECIMAL(\"FF\",16)",
            "ROMAN(499)",
            "ARABIC(\"LVII\")",
            // Onda 2 — Logical + Information + Text escalar + Regex.
            "TRUE()",
            "FALSE()",
            "XOR(TRUE)",
            "IFS(TRUE,1)",
            "SWITCH(1,1,\"a\")",
            "NA()",
            "ISERROR(A1)",
            "ISERR(A1)",
            "ISNA(A1)",
            "ISTEXT(A1)",
            "ISNONTEXT(A1)",
            "ISLOGICAL(A1)",
            "ISEVEN(2)",
            "ISODD(3)",
            "ISREF(A1)",
            "ISFORMULA(A1)",
            "N(1)",
            "T(\"a\")",
            "TYPE(1)",
            "ERROR.TYPE(A1)",
            "SHEETS()",
            "RIGHT(\"ab\",1)",
            "FIND(\"a\",\"ab\")",
            "SEARCH(\"a\",\"ab\",1)",
            "REPLACE(\"abc\",1,1,\"x\")",
            "SUBSTITUTE(\"aa\",\"a\",\"b\",1)",
            "REPT(\"a\",2)",
            "PROPER(\"ab\")",
            "EXACT(\"a\",\"b\")",
            "CHAR(65)",
            "CODE(\"A\")",
            "UNICHAR(66)",
            "UNICODE(\"B\")",
            "CLEAN(\"a\")",
            "FIXED(1234.5,1,TRUE)",
            "DOLLAR(99.888,2)",
            "NUMBERVALUE(\"3.5%\",\".\",\",\")",
            "TEXTBEFORE(\"a-b\",\"-\",1)",
            "TEXTAFTER(\"a-b\",\"-\",-1,1,1,\"none\")",
            "VALUETOTEXT(1,1)",
            "REGEXTEST(\"a\",\"a\",1)",
            "REGEXEXTRACT(\"ab\",\"a\",0)",
            "REGEXREPLACE(\"ab\",\"a\",\"x\",1,1)",
            // Onda 3 — Lookup & Reference escalar.
            "CHOOSE(1,\"a\",\"b\")",
            "HLOOKUP(1,A1:C2,2)",
            "LOOKUP(1,A1:A3,B1:B3)",
            "COLUMN(B1)",
            "COLUMNS(A1:B2)",
            "XMATCH(1,A1:A3,0,-1)",
            "ADDRESS(2,3,1,TRUE,\"S\")",
            "AREAS((A1:A2,B1))",
            "FORMULATEXT(A1)",
            // Onda 4 — condicionais, SUMPRODUCT, estatística descritiva + aliases Compatibility.
            "AVERAGEIF(A1:A3,\">1\")",
            "AVERAGEIFS(A1:A3,B1:B3,1)",
            "MAXIFS(A1:A3,B1:B3,1)",
            "MINIFS(A1:A3,B1:B3,1)",
            "AVERAGEA(A1:A3)",
            "MAXA(A1:A3)",
            "MINA(A1:A3)",
            "SUMPRODUCT(A1:A3,B1:B3)",
            "SUMX2MY2(A1:A3,B1:B3)",
            "SUMX2PY2(A1:A3,B1:B3)",
            "SUMXMY2(A1:A3,B1:B3)",
            "SUBTOTAL(9,A1:A3)",
            "AGGREGATE(9,6,A1:A3)",
            "AGGREGATE(15,6,A1:A3,2)",
            "MEDIAN(A1:A3)",
            "MODE.SNGL(A1:A3)",
            "LARGE(A1:A3,2)",
            "SMALL(A1:A3,2)",
            "RANK.EQ(1,A1:A3)",
            "RANK.AVG(1,A1:A3,1)",
            "PERCENTILE.INC(A1:A3,0.5)",
            "PERCENTILE.EXC(A1:A3,0.5)",
            "PERCENTRANK.INC(A1:A3,1)",
            "PERCENTRANK.EXC(A1:A3,1,3)",
            "QUARTILE.INC(A1:A3,1)",
            "QUARTILE.EXC(A1:A3,1)",
            "TRIMMEAN(A1:A3,0.2)",
            "STDEV.S(A1:A3)",
            "STDEV.P(A1:A3)",
            "STDEVA(A1:A3)",
            "STDEVPA(A1:A3)",
            "VAR.S(A1:A3)",
            "VAR.P(A1:A3)",
            "VARA(A1:A3)",
            "VARPA(A1:A3)",
            "AVEDEV(A1:A3)",
            "DEVSQ(A1:A3)",
            "GEOMEAN(A1:A3)",
            "HARMEAN(A1:A3)",
            "SKEW(A1:A3)",
            "SKEW.P(A1:A3)",
            "KURT(A1:A4)",
            "STANDARDIZE(42,40,1.5)",
            "CORREL(A1:A3,B1:B3)",
            "PEARSON(A1:A3,B1:B3)",
            "COVARIANCE.P(A1:A3,B1:B3)",
            "COVARIANCE.S(A1:A3,B1:B3)",
            "RSQ(A1:A3,B1:B3)",
            "SLOPE(A1:A3,B1:B3)",
            "INTERCEPT(A1:A3,B1:B3)",
            "STEYX(A1:A3,B1:B3)",
            "FORECAST.LINEAR(30,A1:A3,B1:B3)",
            "FISHER(0.75)",
            "FISHERINV(0.972955)",
            "PHI(0.75)",
            "PERMUT(100,3)",
            "PERMUTATIONA(3,2)",
            "PROB(A1:A3,B1:B3,1,3)",
            "MODE(A1:A3)",
            "STDEV(A1:A3)",
            "STDEVP(A1:A3)",
            "VAR(A1:A3)",
            "VARP(A1:A3)",
            "RANK(1,A1:A3)",
            "PERCENTILE(A1:A3,0.5)",
            "PERCENTRANK(A1:A3,1)",
            "QUARTILE(A1:A3,1)",
            "COVAR(A1:A3,B1:B3)",
            "FORECAST(30,A1:A3,B1:B3)",
            // Onda 5 — Date and time.
            "DATE(2026,7,2)",
            "TIME(10,30,0)",
            "DATEVALUE(\"2026-07-02\")",
            "TIMEVALUE(\"10:30\")",
            "YEAR(A1)",
            "MONTH(A1)",
            "DAY(A1)",
            "HOUR(A1)",
            "MINUTE(A1)",
            "SECOND(A1)",
            "DAYS(A1,A2)",
            "DAYS360(A1,A2,TRUE)",
            "EDATE(A1,3)",
            "EOMONTH(A1,3)",
            "WEEKDAY(A1,2)",
            "WEEKNUM(A1,21)",
            "ISOWEEKNUM(A1)",
            "DATEDIF(A1,A2,\"Y\")",
            "YEARFRAC(A1,A2,1)",
            "NETWORKDAYS(A1,A2,A3:A5)",
            "NETWORKDAYS.INTL(A1,A2,\"0000011\",A3:A5)",
            "WORKDAY(A1,10,A3:A5)",
            "WORKDAY.INTL(A1,10,11,A3:A5)",
            // Onda 6 — Financeiras restantes viáveis.
            "SLN(30000,7500,10)",
            "SYD(30000,7500,10,1)",
            "DB(1000000,100000,6,1,7)",
            "DDB(2400,300,10,1)",
            "VDB(2400,300,3650,0,1,2,TRUE)",
            "AMORLINC(2400,A1,A2,300,1,0.15,1)",
            "AMORDEGRC(2400,A1,A2,300,1,0.15,1)",
            "EFFECT(0.0525,4)",
            "NOMINAL(0.053543,4)",
            "MIRR(A1:A3,0.1,0.12)",
            "RRI(96,10000,11000)",
            "PDURATION(0.025,2000,2200)",
            "ISPMT(0.1,1,36,8000000)",
            "CUMIPMT(0.09,360,125000,13,24,0)",
            "CUMPRINC(0.09,360,125000,13,24,0)",
            "FVSCHEDULE(1,A1:A3)",
            "DOLLARDE(1.02,16)",
            "DOLLARFR(1.125,16)",
            "XNPV(0.09,A1:A3,B1:B3)",
            "XIRR(A1:A3,B1:B3,0.1)",
            "ACCRINT(A1,A2,A3,0.1,1000,2,0)",
            "ACCRINTM(A1,A2,0.1,1000,3)",
            "DISC(A1,A2,97.975,100,2)",
            "INTRATE(A1,A2,1000000,1014420,2)",
            "RECEIVED(A1,A2,1000000,0.0575,2)",
            "PRICEDISC(A1,A2,0.0525,100,2)",
            "PRICEMAT(A1,A2,A3,0.061,0.061,2)",
            "YIELDDISC(A1,A2,99.795,100,2)",
            "YIELDMAT(A1,A2,A3,0.0625,100.0123,2)",
            "TBILLEQ(A1,A2,0.0914)",
            "TBILLPRICE(A1,A2,0.09)",
            "TBILLYIELD(A1,A2,98.45)",
            "COUPPCD(A1,A2,2,0)",
            "COUPNCD(A1,A2,2,0)",
            "COUPNUM(A1,A2,2,0)",
            "COUPDAYS(A1,A2,2,0)",
            "COUPDAYBS(A1,A2,2,0)",
            "COUPDAYSNC(A1,A2,2,0)",
            "PRICE(A1,A2,0.0575,0.065,100,2,0)",
            "YIELD(A1,A2,0.0575,95,100,2,0)",
            "DURATION(A1,A2,0.08,0.09,2,1)",
            "MDURATION(A1,A2,0.08,0.09,2,1)",
            "ODDFPRICE(A1,A2,A3,A4,0.0785,0.0625,100,2,0)",
            "ODDFYIELD(A1,A2,A3,A4,0.0575,84.5,100,2,0)",
            "ODDLPRICE(A1,A2,A3,0.0375,0.0405,100,2,0)",
            "ODDLYIELD(A1,A2,A3,0.0375,99.875,100,2,0)",
            // F1 — volatile clock + RNG functions.
            "NOW()",
            "TODAY()",
            "RAND()",
            "RANDBETWEEN(1,6)",
            // Fase 7 — os quatro produtores de array dinâmico. Cada um é um nó de Expression novo
            // (MemoryPackUnion 323-326), e o que os faz escrever NÃO é um arm por função: o FormulaWriter tem
            // UM único `case Function` (FormulaWriter.cs:184) que resolve nome e argumentos por
            // FunctionRegistry.ByType (:437). Então o que esta bateria cobre é a entrada de registro de cada
            // produtor, e um produtor sem entrada estoura NotSupportedException aqui — o que também derrubaria
            // o AnchoredFormulaSupport e o DependencyExtractor, que chegam aos argumentos pelo mesmo acessor.
            // A frase anterior deste comentário afirmava um arm por função, que não existe.
            // As formas com os opcionais entram porque um arm que escreve só os obrigatórios volta a
            // parsear sem erro nenhum, apenas com menos argumentos, e o round trip é o que pega isso.
            "FILTER(A1:A3,B1:B3)",
            "FILTER(A1:A3,B1:B3,\"none\")",
            "SORT(A1:A3)",
            "SORT(A1:B3,2,-1,TRUE)",
            "UNIQUE(A1:A3)",
            "UNIQUE(A1:A3,FALSE,TRUE)",
            "SEQUENCE(3)",
            "SEQUENCE(2,3,10,0.5)",
        ];

        var failures = new List<string>();

        foreach (var formula in corpus)
        {
            var expression = Parse(formula);
            var written = expression.ToFormula(ContextSheet);

            if (!Structure(Parse(written)).SequenceEqual(Structure(expression)))
            {
                failures.Add($"{formula} → {written}");
            }
        }

        await Assert.That(failures).IsEmpty();
    }

    // A DynamicRange (reference-returning endpoints, e.g. INDEX(...) as the ':' left operand) must render
    // through FormulaWriter instead of throwing NotSupportedException, and the written text must re-parse
    // to an equivalent DynamicRange.
    [Test]
    public async Task DynamicRange_RoundTrips()
    {
        var expression = Parse("INDEX($D:$D,2):$D1");
        await Assert.That(expression).IsTypeOf<DynamicRange>();

        var written = expression.ToFormula(ContextSheet);
        await Assert.That(written).Contains(":");

        var reparsed = Parse(written);
        await Assert.That(reparsed).IsTypeOf<DynamicRange>();
    }

    // --- Recursion-depth guard (regression: a programmatically-built tree must throw
    // InvalidOperationException, never overflow the stack — see FormulaWriter.MaxDepth). A parser-produced
    // tree can never exceed the limit (the Parser enforces the same cap at parse time), so these trees are
    // built by hand: a chain of nested UnaryOperation.Negate nodes, deep enough that only a host assembling
    // the AST directly (not through ExpressionParser) could produce it. ---

    private static Expression NestedNegate(int depth)
    {
        Expression expression = new NumberValue(1);

        for (var i = 0; i < depth; i++)
        {
            expression = new UnaryOperation(UnaryOperator.Negate, expression);
        }

        return expression;
    }

    [Test]
    public async Task DeepTree_WithinLimit_Writes()
    {
        // 200 levels: below MaxDepth (256), so this un-parses normally.
        var written = NestedNegate(200).ToFormula(ContextSheet);

        await Assert.That(written).IsEqualTo(new string('-', 200) + "1");
    }

    [Test]
    public async Task DeepTree_ExceedsLimit_ThrowsInvalidOperationException()
    {
        // 300 levels exceed MaxDepth; only a hand-built tree can reach this (the Parser itself caps
        // nesting at parse time), so the writer must guard it independently.
        var expression = NestedNegate(300);

        await Assert
            .That(() => expression.ToFormula(ContextSheet))
            .Throws<InvalidOperationException>();
    }
}
