using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace LabaOOP1
{
    public class TestExcelVisitor : TestExcelBaseVisitor<double>
    {
        //таблиця ідентифікаторів (тут для прикладу)
        //в лабораторній роботі заміните на свою!!!!
        Dictionary<string, double> tableIdentifier = new Dictionary<string, double>();

        public override double VisitCompileUnit(TestExcelParser.CompileUnitContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitNumberExpr(TestExcelParser.NumberExprContext context)
        {
            var result = double.Parse(context.GetText());
            Debug.WriteLine(result);

            return result;
        }

        //IdentifierExpr
        public override double VisitIdentifierExpr(TestExcelParser.IdentifierExprContext context)
        {
            var result = context.GetText();
            double value;
            //видобути значення змінної з таблиці
            if (tableIdentifier.TryGetValue(result.ToString(), out value))
            {
                return value;
            }
            else
            {
                return 0.0;
            }
        }

        public override double VisitParenthesizedExpr(TestExcelParser.ParenthesizedExprContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitExponentialExpr(TestExcelParser.ExponentialExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            Debug.WriteLine("{0} ^ {1}", left, right);
            return System.Math.Pow(left, right);
        }

        public override double VisitAdditiveExpr(TestExcelParser.AdditiveExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            if (context.operatorToken.Type == TestExcelLexer.ADD)
            {
                Debug.WriteLine("{0} + {1}", left, right);
                return left + right;
            }
            else //LabCalculatorLexer.SUBTRACT
            {
                Debug.WriteLine("{0} - {1}", left, right);
                return left - right;
            }
        }

        public override double VisitMultiplicativeExpr(TestExcelParser.MultiplicativeExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            if (context.operatorToken.Type == TestExcelLexer.MULTIPLY)
            {
                Debug.WriteLine("{0} * {1}", left, right);
                return left * right;
            }
            else //LabCalculatorLexer.DIVIDE
            {
                Debug.WriteLine("{0} / {1}", left, right);
                return left / right;
            }
        }
        public override double VisitModDivExpr(TestExcelParser.ModDivExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            if (context.operatorToken.Type == TestExcelLexer.MOD || context.operatorToken.Type == TestExcelLexer.MODMATH)
            {
                Debug.WriteLine("{0} % {1}", left, right);
                return left % right;
            }
            else //LabCalculatorLexer.DIVIDE
            {
                Debug.WriteLine("{0} / {1}", left, right);
                return left / right;
            }
        }

        //public override double VisitModDivExpr(TestExcelParser.ModDivExprContext context)
        //{
        //    var left = WalkLeft(context);
        //    var right = WalkRight(context);

        //    return context.operatorToken.Type == TestExcelLexer.MOD
        //        ? left % right
        //        : (int)left / (int)right;
        //}
        public override double VisitMminExpr(TestExcelParser.MminExprContext context)
        {
            double minValue = Double.PositiveInfinity;

            foreach (var child in context.paramlist.children.OfType<TestExcelParser.ExpressionContext>())
            {
                double childValue = this.Visit(child);
                if (childValue < minValue)
                    minValue = childValue;
            }
            return minValue;
        }
        public override double VisitMinimumExpr(TestExcelParser.MinimumExprContext context)
        {
            double minValue = Double.PositiveInfinity;

            foreach (var child in context.paramlist.children.OfType<TestExcelParser.ExpressionContext>())
            {
                double childValue = this.Visit(child);
                if (childValue < minValue)
                    minValue = childValue;
            }
            return minValue;
        }

        public override double VisitMmaxExpr(TestExcelParser.MmaxExprContext context)
        {
            double maxValue = double.NegativeInfinity;

            foreach (var child in context.paramlist.children.OfType<TestExcelParser.ExpressionContext>())
            {
                double childValue = Visit(child);
                if (childValue > maxValue)
                    maxValue = childValue;
            }
            return maxValue;
        }
        public override double VisitMaximumExpr(TestExcelParser.MaximumExprContext context)
        {
            double maxValue = Double.NegativeInfinity;

            foreach (var child in context.paramlist.children.OfType<TestExcelParser.ExpressionContext>())
            {
                double childValue = Visit(child);
                if (childValue > maxValue)
                    maxValue = childValue;
            }
            return maxValue;
        }

        public override double VisitPowExpr(TestExcelParser.PowExprContext context)
        {
            double maxValue = Double.NegativeInfinity;
            double[] num = { 0, 1 };
            int i = 0;
            foreach (var child in context.paramlist.children.OfType<TestExcelParser.ExpressionContext>())
            {
                double childValue = this.Visit(child);
                if (childValue > maxValue)
                    maxValue = childValue;
                num[i] = childValue;
                ++i;
            }
            return Math.Pow(num[0], num[1]);
        }


        private double WalkLeft(TestExcelParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<TestExcelParser.ExpressionContext>(0));
        }

        private double WalkRight(TestExcelParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<TestExcelParser.ExpressionContext>(1));
        }
    }
}
