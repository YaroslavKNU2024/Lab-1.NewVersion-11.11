using Antlr4.Runtime;
using LabaOOP1;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestLab1
{
    [TestClass]
    public class UnitTest1
    {
        public static TestExcelLexer Lexer(AntlrInputStream stream) => new TestExcelLexer(stream);
        [TestMethod]
        public void TestMethod1()
        {
            string s = "1+2*75-5";
            Calculator calc = new Calculator(s);

            var lexer = Lexer(new AntlrInputStream(s));
            Assert.AreEqual(lexer.ParseInfo, calc.lexer.ParseInfo);
        }
        [TestMethod]
        public void TestMethod2()
        {
            string s = "1+2*75-5";
            string expected = "146";
            string actual = new Calculator(s).Evaluate(s).ToString();

            Assert.AreEqual(expected, actual);
        }
    }
}
