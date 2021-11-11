
namespace LabaOOP1
{
    public class CellElement
    {
        public int IndRow;
        public int IndCol;
        public string Name;
        public string Expression;
        public string Value;

        public CellElement() { }
        public CellElement(int row, int col, string expression, string value)
        {
            IndRow = row;
            IndCol = col;
            Name = "@" + "R" +(IndRow+1).ToString() + "C" + (IndCol+1).ToString();
            Expression = expression;
            Value = value;
        }

    }
}
