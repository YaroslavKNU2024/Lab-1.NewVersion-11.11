using Antlr4.Runtime;
using LabaOOP1;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestLab3
{
    [TestClass]
    public class UnitTest1
    {
        string s = "34/2*5";
        string expected = "85";
        public static TestExcelLexer Lexer(AntlrInputStream stream) => new TestExcelLexer(stream);
        public static CommonTokenStream Tokenizer(TestExcelLexer lexer) => new CommonTokenStream(lexer);
        public static TestExcelParser Parser(CommonTokenStream token) => new TestExcelParser(token);
        [TestMethod]
        public void TestMethod1()
        {
            var lexer = new TestExcelLexer(new AntlrInputStream(s));
            var tokens = Tokenizer(lexer);
            var ps = Parser(tokens);

            Calculator calc = new Calculator(s);
            calc.Evaluate(s);
            Assert.AreEqual(calc.parser.ParseInfo, ps.ParseInfo);
        }
        [TestMethod]
        public void TestMethod2()
        {
            Calculator calc = new Calculator(s);
            string actual = calc.Evaluate(s).ToString();
            Assert.AreEqual(expected, actual);
        }
    }
}
