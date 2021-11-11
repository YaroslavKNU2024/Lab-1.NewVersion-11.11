using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace LabaOOP1
{
    public static class Checker
    {
        private static bool ThrowError()
        {
            MessageBox.Show("Parenthesis error!");
            _lparen = 0;
            _rparen = 0;
            return false;
        }
        // function to check correctness
        private static readonly HashSet<char> _digits = new HashSet<char>() { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
        private static int _lparen = 0, _rparen = 0;
        private static bool CheckParentNumberError(char x)
        {
            if (x == '(')
                _lparen++;
            else if (x == ')')
            {
                _rparen++;
                if (_lparen < _rparen)
                {
                    ThrowError();
                    return false;
                }
            }
            return true;
        }
        public static bool CheckCorrect(string expr)
        {
            for (int i = 0, len = expr.Length; i < len; ++i)
            {
                if (expr[i] == '(' || expr[i] == ')')
                {
                    if (!CheckParentNumberError(expr[i]))
                        return false;
                }
                else if (_digits.Contains(expr[i]) && (i != 0 && i != len - 1 && (expr[i + 1] == '(' || expr[i - 1] == ')')) || (i == 0 && expr[0] == ')') || (i == len - 1 && expr[i] == '('))
                {
                    ThrowError();
                    return false;
                }
            }
            if (_lparen != _rparen) 
            {
                ThrowError();
                return false;
            }
            _lparen = _rparen = 0;
            return true;
        }
    }
}
