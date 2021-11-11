using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace LabaOOP1
{
    public partial class Saver : Form
    {
        public Saver()
        {
            InitializeComponent();
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void SetFileName_Click(object sender, EventArgs e)
        {
            FileName.path = textBox1.Text;
            MessageBox.Show("Ім'я записано вдало");
            //MessageBox.Show(FileName.path);
            Hide();
            string json = JsonConvert.SerializeObject(FileName.Cell.ToArray());
            File.WriteAllText(FileName.path, json);
        }
    }
}
