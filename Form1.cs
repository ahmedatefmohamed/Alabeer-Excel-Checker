using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace AlabeerExcelCheckerv2
{
    public partial class Form1 : Form
    {
        string filePath = string.Empty;
        OpenFileDialog file = new OpenFileDialog();
        string fileExtention = string.Empty;

        Excel.Application excelApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range, // All sheet range.
                    sortedRange, // range for column should sort to.
                    hiddenRange; // range for columns want to hide.

        public void sharedProcess()
        {
            filePath = file.FileName; //get the path of the file ex.(D\sheet.xlsx) 
            fileExtention = Path.GetExtension(filePath); //get the file extension ex.(.xlsx)
            String fileName = Path.GetFileNameWithoutExtension(filePath);

            excelApp = new Excel.Application();
            xlWorkBook = excelApp.Workbooks.Open(filePath); // open excel file.
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); // get and read the sheet number 1.

            int rowsCount = xlWorkSheet.UsedRange.Rows.Count - 2; // except header row {13} and total row @the last bottom.
            int columnsCount = xlWorkSheet.UsedRange.Columns.Count; // get number of columns.

            /** MessageBox.Show("Rows count: " + rowsCount); **/

            range = xlWorkSheet.Range[$"A14:A{rowsCount}", "AF13"];
            sortedRange = xlWorkSheet.Range[$"D14:D{rowsCount}"];
            hiddenRange = xlWorkSheet.Range[$"G14:G{rowsCount}", "Z13"];

            // Column index will be sum (Net {27}, VAT {28} and Billed {29})
            int[] arr = new int[] { 27, 28, 29 };

            hiddenRange.EntireColumn.Hidden = true; // hide unwanted range columns.

            // Sort Insured Company Name Column...
            sortedRange.Select();
            xlWorkSheet.Sort.SortFields.Clear();
            xlWorkSheet.EnableAutoFilter = true;
            xlWorkSheet.AutoFilter.Sort.SortFields.Add(sortedRange, Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlAscending, System.Type.Missing, Excel.XlSortDataOption.xlSortNormal);
            xlWorkSheet.AutoFilter.ApplyFilter();
            xlWorkSheet.Sort.SetRange(sortedRange);
            xlWorkSheet.Sort.Header = Excel.XlYesNoGuess.xlGuess;
            xlWorkSheet.Sort.MatchCase = false;
            xlWorkSheet.Sort.Orientation = (XlSortOrientation)Excel.Constants.xlTopToBottom;
            xlWorkSheet.Sort.SortMethod = Excel.XlSortMethod.xlPinYin;
            xlWorkSheet.Sort.Apply();

            // Grouped by Company Name
            range.Subtotal(4, Excel.XlConsolidationFunction.xlSum, new int[] { 27, 28, 29 }, true, true, Excel.XlSummaryRow.xlSummaryBelow); ;
            range.Activate();
        }

        public void groupExcel(bool isExport) {

            if(isExport == true)
            {
                try
                {
                    sharedProcess();

                    // Open Excel sheet and enable visibility.
                    excelApp.Visible = true;

                    xlWorkBook.Close(true, null, null);
                    excelApp.Quit();

                    releaseObject(excelApp);
                    releaseObject(xlWorkBook);
                    releaseObject(xlWorkSheet);
                }
                catch (Exception ex)
                {
                    // MessageBox.Show(ex.Message.ToString() + " Error!");
                }
            } 
            else
            {
                String fileName = Path.GetFileNameWithoutExtension(filePath);
                if (listBox1.Items.Contains(fileName) == true)
                {
                    MessageBox.Show("Already done!");
                }
                else
                {
                    try
                    {
                        sharedProcess();
                        listBox1.Items.Add(fileName);
                        txtFileCount.Text = listBox1.Items.Count.ToString();

                        // Open Excel sheet and enable visibility.
                        excelApp.Visible = true;

                        xlWorkBook.Close(true, null, null);
                        excelApp.Quit();

                        releaseObject(excelApp);
                        releaseObject(xlWorkBook);
                        releaseObject(xlWorkSheet);
                    }
                    catch (Exception ex)
                    {
                        // MessageBox.Show(ex.Message.ToString() + " Error!");
                    }
                }
            }

            
        }

        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Firing Upload event
        /// </summary>
        private void btn_upload(object sender, EventArgs e)
        {
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filePath = file.FileName; //get the path of the file ex.(D\sheet.xlsx) 
                fileExtention = Path.GetExtension(filePath); //get the file extension ex.(.xlsx)

                if (fileExtention.CompareTo(".xls") == 0 || fileExtention.CompareTo(".xlsx") == 0)
                {
                    try { txtProAttach.Text = filePath; }
                    catch (Exception ex) { MessageBox.Show(ex.Message.ToString()); }
                }
                else
                {
                    //dispaly a custom messageBox to show error  
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                }
            }
        }

        /// <summary>
        /// Firing Group event
        /// </summary>
        private void btn_group(object sender, EventArgs e)
        {
            groupExcel(false);
        }

        /// <summary>
        /// Firing Export event
        /// </summary>
        private void btn_export(object sender, EventArgs e)
        {
            groupExcel(true);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}
