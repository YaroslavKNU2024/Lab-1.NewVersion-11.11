using System;
using System.Windows.Forms;

namespace LabaOOP1
{
    public partial class Loader : Form
    {
        public Loader()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FileName.path = textBox1.Text;
            MessageBox.Show("Ім'я записано вдало");
            Hide();
        }
    }
}
