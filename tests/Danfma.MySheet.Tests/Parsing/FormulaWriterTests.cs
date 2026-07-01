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
    [Arguments("SUM((A1:A3,C1:C3))")] // união de ranges
    [Arguments("IF(A1>0,\"yes\",\"no\")")]
    [Arguments("IFERROR(1/0,42)")]
    [Arguments("VLOOKUP(A1,B1:C10,2,FALSE)")]
    [Arguments("COUNTIF(A1:A10,\">5\")")]
    [Arguments("LET(x,1,y,2,x+y)")]
    [Arguments("SHEET()")]
    [Arguments("MYFN(1,A1)")] // função custom (FunctionCall)
    [Arguments("ROUND(1.234,2)")]
    [Arguments("Sheet2!A1+1")]
    [Arguments("Sheet2!A1:B2")]
    [Arguments("'My Sheet'!A1")]
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
    public async Task NormalizesEquivalentText(string input, string canonical)
    {
        var expression = Parse(input);
        var written = expression.ToFormula(ContextSheet);

        await Assert.That(written).IsEqualTo(canonical);
        await Assert.That(Structure(Parse(written)).SequenceEqual(Structure(expression))).IsTrue();
    }

    // Uma chamada mínima válida por função built-in do Parser: o texto re-parseado tem que produzir
    // exatamente a mesma árvore (garante que o mapa nó→nome cobre todas as funções).
    [Test]
    public async Task AllBuiltInFunctions_RoundTripStructurally()
    {
        string[] corpus =
        [
            "SUM(1)", "AVERAGE(1)", "MIN(1)", "MAX(1)", "COUNT(1)",
            "IF(TRUE,1,2)", "AND(TRUE)", "OR(TRUE)", "NOT(TRUE)", "IFERROR(1,2)", "IFNA(1,2)",
            "INT(1.5)", "ROUND(1.234,2)", "ROUNDUP(1.2,0)", "ABS(1)",
            "ISNUMBER(1)", "ISBLANK(A1)",
            "UPPER(\"a\")", "LOWER(\"A\")", "TRIM(\" a \")", "LEN(\"a\")", "LEFT(\"ab\",1)",
            "MID(\"abc\",2,1)", "VALUE(\"1\")", "TEXT(1,\"0\")",
            "CONCAT(\"a\")", "CONCATENATE(\"a\",\"b\")", "TEXTJOIN(\",\",TRUE,\"a\")",
            "COUNTA(A1:A2)", "COUNTBLANK(A1:A2)", "COUNTIF(A1:A2,1)", "COUNTIFS(A1:A2,1)",
            "SUMIF(A1:A2,1)", "SUMIFS(A1:A2,B1:B2,1)",
            "ROWS(A1:A2)", "ROW()", "MATCH(1,A1:A3)", "INDEX(A1:B2,1,1)",
            "VLOOKUP(1,A1:B2,2)", "XLOOKUP(1,A1:A2,B1:B2)", "OFFSET(A1,1,1)",
            "LET(x,1,x)", "SHEET()",
            "PMT(0.05,10,1000)", "PV(0.05,10,100)", "FV(0.05,10,100)", "NPER(0.05,100,1000)",
            "IPMT(0.05,1,10,1000)", "PPMT(0.05,1,10,1000)", "NPV(0.1,A1:A3)",
            "RATE(10,100,1000)", "IRR(A1:A3)",
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
}
