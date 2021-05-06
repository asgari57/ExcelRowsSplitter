using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;        //Microsoft Excel 14 object in references-> COM tab

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public static int rowCount;
        public static int  colCount;
        public int sec = 0;
        public int min = 0;
        public int hr = 0;

        public Form1()
        {
            InitializeComponent();
            



        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog()==DialogResult.OK)

            txtSave.Text = folderBrowserDialog1.SelectedPath;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // openFileDialog1.ShowDialog();
            // txtFile.Text = openFileDialog1.
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtFile.Text = openFileDialog1.FileName;


            }



        }


       
       

       
        private void button4_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(txtFile.Text);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            rowCount = xlRange.Rows.Count;
            colCount = xlRange.Columns.Count;



            nRows.Text = rowCount.ToString();
            nCoulmn.Text = colCount.ToString();

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            timer1.Enabled = true;

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(txtFile.Text);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            rowCount = xlRange.Rows.Count;
            colCount = xlRange.Columns.Count;
            int number;
          
            int c = 0;
            int cc = 0;
            bool isParsable = Int32.TryParse(txtRows.Text, out number);
            if (isParsable)
                c = number;
            object misValue = System.Reflection.Missing.Value;

            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            xlWorkBook2 = xlApp.Workbooks.Add(misValue);
            xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(1);

            int nfile = 0;


            progressBar1.Maximum = rowCount;
            progressBar1.Step = 1;

          //  int r = 0;
            for (int x = 1; x <= colCount; x++)
            {
                xlWorkSheet2.Cells[1, x] = xlWorksheet.Cells[1, x];

            }

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    if (cc <= c && i<=c )
                    {
                        xlWorkSheet2.Cells[i-cc, j] = xlWorksheet.Cells[i, j];
                        
                        if (i-cc == 1 && checkBox1.Checked) { 
                        xlWorkSheet2.Cells[1, j] = xlWorksheet.Cells[1,j];
                        }
                    }
                }
                if (i == c && c <= rowCount)
                {
                    // r = 0;
                    cc = c;
                    c = c + c;

                    
                    
                    nfile += 1;
                    xlWorkBook2.SaveAs(txtSave.Text + "\\split" + nfile + ".xlsx", Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook2.Close(true, misValue, misValue);
                    xlWorkBook2 = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(1);

                    
                    
                }

                
                progressBar1.PerformStep();
                
            }

            nfile += 1;
            xlWorkBook2.SaveAs(txtSave.Text + "\\split" + nfile + ".xlsx", Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook2.Close(true, misValue, misValue);
            Marshal.ReleaseComObject(xlWorkSheet2);
            Marshal.ReleaseComObject(xlWorkBook2);








            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);




            MessageBox.Show("Finished!", "Process information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            progressBar1.Value = 0;
          

        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void lblTimer_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            //sec++;
            //if (sec == 60)
            //{
            //    min++;
            //    sec = 0;
            //}
            //if (min == 60)
            //{
            //    hr++;
            //    min = 0;
            //    sec = 0;
            //}


            //lblTimer.Text = hr.ToString() + ":" + min.ToString() + ":" + sec.ToString();



        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkLabel1.LinkVisited = true;
            System.Diagnostics.Process.Start("https://www.asiahkz.com");

        }
    }
}
