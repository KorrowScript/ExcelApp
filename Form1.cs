using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDesignApp
{
    public partial class Form1 : Form
    {
        private List<string> valuesList = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Snatch input values from those text boxes
            string param1 = textBox1.Text;
            valuesList.Add(param1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Summon an Excel application instance from the software void
            Excel.Application excelApp = new Excel.Application();
            if (excelApp == null)
            {
                MessageBox.Show("Oops! Excel isn't here. Is it installed?");
                return;
            }

            // Conjure a new Excel workbook out of thin air
            Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            worksheet.Name = "Magic Data";

            // Place column headers like a boss
            worksheet.Cells[1, 1] = "Index";
            worksheet.Cells[1, 2] = "Value";

            // Insert the precious data from the list into the cells
            for (int i = 1; i < valuesList.Count; i++)
            {
                worksheet.Cells[i + 2, 1] = i; // Index column
                worksheet.Cells[i + 2, 2] = valuesList[i]; // Value column
            }

            // Tell those columns to shape up and auto-fit!
            worksheet.Columns.AutoFit();

            // Save the magical workbook to the specified path
            string filePath = @"C:\Users\klagi\Desktop\file.xlsx";
            workbook.SaveAs(filePath);
            workbook.Close();
            excelApp.Quit();

            MessageBox.Show("Behold! Your Excel file has been created!");
        }

    }
}