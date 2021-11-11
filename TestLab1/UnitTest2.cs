using Microsoft.VisualStudio.TestTools.UnitTesting;
using LabaOOP1;
using Antlr4.Runtime;

namespace TestLab2
{
    [TestClass]
    public class UnitTest1
    {
        string s = "mmax(2, 3)*8^4";
        string expected = "12288";
        public static TestExcelLexer Lexer(AntlrInputStream stream) => new TestExcelLexer(stream);
        public static CommonTokenStream Tokenizer(TestExcelLexer lexer) => new CommonTokenStream(lexer);
        [TestMethod]
        public void TestMethod1()
        {
            var lexer = new TestExcelLexer(new AntlrInputStream(s));
            var tokens = Tokenizer(lexer);
            Calculator calc = new Calculator(s);
            calc.Evaluate(s);

            Assert.AreEqual(calc.tokens.ToString(), tokens.ToString());
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
