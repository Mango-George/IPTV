using ExcelLoader.Scripts;
using System;
using System.Windows.Forms;

namespace ExcelLoader
{
    public partial class Form1 : Form
    {
        private ExcelHandler excelHandler;
        public Form1()
        {
            excelHandler = new ExcelHandler();
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            excelHandler.Export();
        }
    }
}
