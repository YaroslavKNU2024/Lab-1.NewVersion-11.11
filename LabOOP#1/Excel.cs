using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LabaOOP1
{
    public class Calculator
    {
        public Calculator(string exp) {
            lexer = Lexer(new AntlrInputStream(exp));
        }
        public Calculator() {}
        public TestExcelLexer lexer;
        public CommonTokenStream tokens;
        public TestExcelParser parser;

        public TestExcelLexer Lexer(AntlrInputStream stream) => new TestExcelLexer(stream);
        public CommonTokenStream Tokenizer(TestExcelLexer lexer) => new CommonTokenStream(lexer);
        public TestExcelParser Parser(CommonTokenStream token) => new TestExcelParser(token);

        public double Evaluate(string expression)
        {
            var v = new AntlrInputStream(expression);
            lexer = Lexer(v);
            //var lexer = new TestExcelLexer(new AntlrInputStream(expression));

            lexer.RemoveErrorListeners();
            lexer.AddErrorListener(new ThrowExceptionErrorListener());

            //var tokens = new CommonTokenStream(lexer);
            //var parser = new TestExcelParser(tokens);
            tokens = Tokenizer(lexer);
            parser = Parser(tokens);

            var tree = parser.compileUnit();
            var visitor = new TestExcelVisitor();
            double dd = visitor.Visit(tree);
            return dd;
        }
    }
}
