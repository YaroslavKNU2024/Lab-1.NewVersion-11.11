using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace LabaOOP1
{
    public partial class MiniExcel : Form
    {
        private int _row = 10;
        private int _col = 10;
        public List<CellElement> _cells = new List<CellElement>();
        public MiniExcel()
        {
            InitializeComponent();
            WindowState = FormWindowState.Maximized;
            InitializeDataGridView();
        }
        private void InitializeDataGridView()
        {
            dgv.AllowUserToAddRows = false;
            dgv.ColumnCount = _col;
            dgv.RowCount = _row;
            FillHeaders();
            dgv.AutoResizeRows();
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.RowHeadersWidth = 100;
        }

        private void FillHeaders()
        {
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                col.HeaderText = "C" + (col.Index + 1);
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = "R" + (row.Index + 1);
            }
        }

        // get values from other cells
        private string ResolveAddresses(string s)
        {
            string x = "";
            int i = 0, len = s.Length;
            while (i < len)
            {
                if (s[i] == '@')
                {
                    Tuple<int, int> t = GetIndecesFromName(s.Substring(i + 1, len - i - 1));
                    try
                    {
                        if (dgv.Rows[t.Item1].Cells[t.Item2].Value.ToString() == "") return "НЕ ІСНУЄ КЛІТИНКИ";
                        if (s == dgv.Rows[t.Item1].Cells[t.Item2].Value.ToString()) return "ПОМИЛКА! РЕКУРСІЯ ВИЯВЛЕНА!";
                        x += ResolveAddresses(dgv.Rows[t.Item1].Cells[t.Item2].Value.ToString());
                    } catch (Exception ex)
                    {
                        if (ex is NullReferenceException)
                            return "НЕ ІСНУЄ КЛІТИНКИ";
                        if (ex is ArgumentOutOfRangeException || ex is IndexOutOfRangeException)
                            return "КРИТИЧНА ПОМИЛКА";
                    }
                    i += 3 + (t.Item1+1).ToString().Length + (t.Item2+1).ToString().Length;
                }
                else {
                    x += s[i];
                    ++i;
                }
            }
            return x;
        }

        // get indeces from the name of a cell (ex. R2C1)
        private Tuple<int, int> GetIndecesFromName(string s)
        {
            string slen = "";
            foreach (char n in s)
                if (n <= '9' && n >= '0')
                    slen += n;
                else if (n == 'C')
                    break;
            int RowIndex = int.Parse(slen);
            slen = "";
            bool col = false;
            foreach (char n in s)
                if (n == 'C') col = true;
                else if (col)
                {
                    if (n <= '9' && n >= '0')
                        slen += n;
                    else break;
                }
            int ColumnIndex = int.Parse(slen);
            return new Tuple<int, int>(RowIndex-1, ColumnIndex-1);
        }

        // updating values in cells
        private void UpdateAllCells()
        {
            for (int i = 0; i < _cells.Count; ++i)
            {
                string expr = _cells[i].Expression;
                string ValueOfCell = ResolveAddresses(expr);
                _cells[i].Value = ValueOfCell;
            }
            for (int i = 0; i < _cells.Count; ++i)
            {
                Tuple<int, int> t = GetIndecesFromName(_cells[i].Name);
                if (_cells[i].Value.ToString() == "НЕ ІСНУЄ КЛІТИНКИ")
                {
                    if (t.Item1 == dgv.CurrentCell.RowIndex && t.Item2 == dgv.CurrentCell.ColumnIndex)
                    {
                        MessageBox.Show("Помилка при обробці клітинки! Перевірте існування клітинки та правильність вводу");
                        _cells.RemoveAt(i);
                        --i;
                    }
                    continue;
                }
                if (_cells[i].Value.ToString() == "КРИТИЧНА ПОМИЛКА")
                {
                    if (t.Item1 == dgv.CurrentCell.RowIndex && t.Item2 == dgv.CurrentCell.ColumnIndex)
                    {
                        MessageBox.Show("Помилка при обробці клітинки! Перевірте правильність введених параметрів");
                        _cells.RemoveAt(i);
                        --i;
                    }
                    continue;
                }
                if (_cells[i].Value.ToString() == "ПОМИЛКА! РЕКУРСІЯ ВИЯВЛЕНА!")
                {
                    MessageBox.Show("ПОМИЛКА! РЕКУРСІЯ ВИЯВЛЕНА!\nКЛІТИНКУ НЕ БУЛО СТВОРЕНО!");
                    _cells.RemoveAt(i);
                    --i;
                    continue;
                }
                Calculator calc = new Calculator();
                var res = calc.Evaluate(_cells[i].Value.ToString());
                if (res.ToString() == "∞" || res.ToString() == "-∞")
                    dgv.Rows[t.Item1].Cells[t.Item2].Value = "ПОМИЛКА! ДІЛЕННЯ НА НУЛЬ!";
                else
                    dgv.Rows[t.Item1].Cells[t.Item2].Value = res;
            }
        }
        // hide Values & show expressions
        private void ShowExpressions()
        {
            for (int i = 0; i < _cells.Count; ++i)
            {
                Tuple<int, int> t = GetIndecesFromName(_cells[i].Name);
                dgv.Rows[t.Item1].Cells[t.Item2].Value = _cells[i].Expression;
            }
        }

        // change a selected cell
        private void ApplyButton_Click(object sender, EventArgs e)
        {
            if (dgv.CurrentCell == null) return;
            int rowIndex = dgv.CurrentCell.RowIndex;
            int columnIndex = dgv.CurrentCell.ColumnIndex;
            if (dgv.CurrentCell.Value == null)
            {
                if (TextBox.Text == string.Empty) return;
                dgv.CurrentCell.Value = TextBox.Text;
                TextBox.Text = string.Empty;
            }
            CellElement Cell = new CellElement(rowIndex, columnIndex, dgv.CurrentCell.Value.ToString(), "to be implemented");

            string expr = dgv.CurrentCell.Value.ToString();
            if (!Checker.CheckCorrect(expr))
            {
                dgv.CurrentCell.Value = "ПОМИЛКА ДУЖОК!!!";
                return;
            }
            string ValueOfCell = ResolveAddresses(expr);
            Cell.Value = ValueOfCell;

            var match = _cells.FirstOrDefault(c => _cells.Contains(Cell));
            if (match == null)
                _cells.Add(Cell);
            dgv.CurrentCell.Value = Cell.Value.ToString();
            int _numOfIterations = _cells.Count * (_cells.Count - 1);
            for (int k =0; k < _numOfIterations+1; ++k)
                UpdateAllCells();
        }

        // add row / column
        private void AddRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _row++;
            dgv.RowCount = _row;
            FillHeaders();
        }

        private void AddColumnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _col++;
            dgv.ColumnCount = _col;
            FillHeaders();
        }


        private void ShowExpressionsToolStripMenuItem_Click(object sender, EventArgs e) => ShowExpressions();
        

        private void ShowValuesToolStripMenuItem_Click(object sender, EventArgs e) => UpdateAllCells();

        readonly List<CellElement> _clean = new List<CellElement>();
        private void MakeNull()
        {
            for (int i = 0; i < _row; ++i)
                for (int j = 0; j < _col; ++j)
                    dgv.Rows[i].Cells[j].Value = "";
            _cells = _clean;
        }
        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MakeNull();
            using (Loader newProjectForm = new Loader()) {
                if (newProjectForm.ShowDialog() == DialogResult.OK){}
            }
            try
            {
                string x = File.ReadAllText(FileName.path);
                List<CellElement> ca = JsonConvert.DeserializeObject<List<CellElement>>(x);
            
                foreach (CellElement elem in ca)
                {
                    if (elem.IndRow > _row)
                    {
                        while(elem.IndRow >= _row)
                            AddRowToolStripMenuItem_Click(sender, e);
                    }
                    if (elem.IndCol >= _col)
                    {
                        while(elem.IndCol> _col)
                            AddColumnToolStripMenuItem_Click(sender, e);
                    }
                    CellElement _newCellElement = new CellElement(elem.IndRow, elem.IndCol, elem.Expression, elem.Value);
                    _cells.Add(_newCellElement);
                }
                for (int k = 0, numberOfIterations = _cells.Count+1; k < numberOfIterations; ++k)
                    UpdateAllCells();
            }
            catch (FileNotFoundException) { MessageBox.Show("Не існує файлу!"); return; }
        }

        public string ShowMe { get; set; }
        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileName.Cell = _cells;
            Saver sv = new Saver();
            sv.Show();
        }

        private void DeleteRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (--_row < 0)
                MessageBox.Show("Неможливо видалити рядок");
            else
                dgv.Rows.RemoveAt(_row);
            for (int i = 0; i < _cells.Count; ++i)
            {
                Tuple<int, int> indecesFromName= GetIndecesFromName(_cells[i].Name);
                Tuple<int, int> indecesFromExpression = new Tuple<int, int>(-1, -1);
                if (_cells[i].Expression.Length > 1 && _cells[i].Expression[0] == '@')
                    indecesFromExpression = GetIndecesFromName(_cells[i].Expression);
                if (indecesFromExpression.Item1 >= _row && indecesFromName.Item1 < _row)
                {
                    dgv.Rows[indecesFromName.Item1].Cells[indecesFromName.Item2].Value = "";
                    _cells.RemoveAt(i);
                    --i;
                } else if (indecesFromName.Item1 >= _row)
                {
                    _cells.RemoveAt(i);
                    --i;
                }
            }
            UpdateAllCells();
        }

        private void DeleteColumnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (--_col < 0)
                MessageBox.Show("Неможливо видалити стовбець");
            else
                dgv.Columns.RemoveAt(_col);
            for (int i = 0; i < _cells.Count; ++i)
            {
                Tuple<int, int> indecesFromName = GetIndecesFromName(_cells[i].Name);
                Tuple<int, int> indecesFromExpression = new Tuple<int, int>(-1, -1);
                if (_cells[i].Expression.Length > 1 && _cells[i].Expression[0] == '@')
                    indecesFromExpression = GetIndecesFromName(_cells[i].Expression);
                if (indecesFromExpression.Item2 >= _col && indecesFromName.Item2 < _col)
                {
                    dgv.Rows[indecesFromName.Item1].Cells[indecesFromName.Item2].Value = "";
                    _cells.RemoveAt(i);
                    --i;
                }
                else if (indecesFromName.Item2 >= _col)
                {
                    _cells.RemoveAt(i);
                    --i;
                }
            }
            UpdateAllCells();
        }

        private void HelpToolStripMenuItem_Click(object sender, EventArgs e)
        {

            string text = "Метою даної програми було створення міні-версії Екселя. Користувач може вводити в клітинки математичні вирази, такі як +, -, ^, *, /, mod, div, pow, mmax, mmin, числа, посилання на інші клітинки тощо.\n" +
                " Після введення необхідного виразу необхідно натиснути кнопку \"Обчислити\" для того, щоб операція була виконана. \n" +
                "У \"Вигляді клітинок\" можна також переглянути вирази / значення" +
                "введені користувачем\nДля збереження / відкриття / занулювання таблиці необхідно вибрати відповідну команду в menu-strip \"Обробка таблиць\"\n" +
                "У strip \"Форматування таблиці\" можна додати / видалити рядок\n" +
                "Замість комірки, можна також вводити програму в тектбоксі, зліва від кнопки \"Обчислити\". \n" +
                "До відповідної виділеної клітинки тоді будуть застосовані введені  зміни. Це особливо зручно для довгих математичних виразів";
            MessageBox.Show(text);
        }

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string text = "Дана програма є міні-версією Екселя, яка імітує весь функціонал обробки та роботи з таблицями. Творець: Ярослав Сакаль К-24";
            MessageBox.Show(text);
        }

        private void NewTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            while(_row != 10)
            {
                DeleteRowToolStripMenuItem_Click(sender, e);
            }
            while(_col != 10)
            {
                DeleteColumnToolStripMenuItem_Click(sender, e);
            }
            MakeNull();
        }
    }
}
