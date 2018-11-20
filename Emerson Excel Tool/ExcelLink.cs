using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

namespace Emerson_Excel_Tool
{
    public partial class ToolForm : Form
    {
        #region Excel Linking

        //"http://aka.ms/dotnet-get-started-desktop");


        public class ExcelProcess
        {

            private Excel.Application oXL;
            private Excel._Workbook oWB;
            private Excel._Worksheet oSheet;
            private Excel.Range oRng;

            public ExcelProcess()
            {
                ExcelLauncher(out oXL, out oWB, out oSheet);
                oXL.Visible = true;
                oWB = oXL.ActiveWorkbook;
                oSheet = (Excel.Worksheet)oWB.ActiveSheet;
            }
            


            public void Close()
            {
                oXL.Visible = true;
                bool SaveChanges;
                oWB.Close(SaveChanges = false,Type.Missing , Type.Missing);
                oXL.Quit();
            }

            public void Launch(List<string> testFileList)
            {
                if (testFileList.Count < 1)
                {
                    MessageBox.Show("No data selected to import!", "Error");
                }
                else
                {
                    try
                    {


                        //creates sheets labeled 1-4
                        //for (int i = 1; i < 5; i++)
                        //{
                        //    int count = oWB.Worksheets.Count;
                        //    Excel.Worksheet addedSheet = oWB.Worksheets.Add(Type.Missing,
                        //            oWB.Worksheets[count], Type.Missing, Type.Missing);
                        //    addedSheet.Name = i.ToString();
                        //}

                        /*By doing the following, you will create a named range(Transactions) on the oSheet staring at cell A1 and finising at cell C3

                Range namedRange = oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[3, 3]];
                oSheet.Names.Add("Transactions", newRange);
                        namedRange.Name = "Transactions";*/



                        int columnNumb = 2 * testFileList.Count;
                        for (int i = 0; i < testFileList.Count; i++)
                        {

                            int j = (2 * i) + 1;
                            int k = (2 * i) + 2;
                            //Add table headers going cell by cell.
                            oSheet.Cells[1, j] = "Frequency" + (1 + i);
                            oSheet.Cells[1, k] = "Response" + (1 + i);
                        }

                        string letterVal = IndexToColumn(columnNumb);


                        //FORMATTING
                        //Format headers as bold, vertical alignment = center.
                        oSheet.get_Range("A1", letterVal + "1").Font.Bold = true;
                        oSheet.get_Range("A1", letterVal + "1").VerticalAlignment =
                            Excel.XlVAlign.xlVAlignCenter;
                        oSheet.get_Range("A1", letterVal + "1").ColumnWidth = 18;



                        //attempt to fill excel with DataTables object.
                        DataTable dt;
                        string[,] results;

                        List<DataSet_Processing> listOfDataSets = new List<DataSet_Processing>();

                        for (int i = 0; i < testFileList.Count; i++)
                        {
                            listOfDataSets.Add(new DataSet_Processing { tableName = "File #" + (i + 1) });
                        }
                        int totalDataSetsLoaded = listOfDataSets.Count;
                        int _filecounter = 0;
                        for (int i = 0; i < (2 * listOfDataSets.Count); i = i + 2)
                        {
                            // There's something fucky here...
                            listOfDataSets[_filecounter].tableFileLocation = testFileList[_filecounter];
                            listOfDataSets[_filecounter].tableName = Path.GetFileName(testFileList[_filecounter]);
                            listOfDataSets[_filecounter].GetTableData(out dt, out results);
                            oSheet.get_Range(IndexToColumn(i + 1) + dt.Columns.Count, IndexToColumn(i + 2) + dt.Rows.Count).Value2 = results;
                            _filecounter++;

                        }

                        oRng = oSheet.Range["A1:D489"];
                        object[,] transposedRange = (object[,])oXL.WorksheetFunction.Transpose(oRng.Value2);
                        Excel.Worksheet oSheet2;
                        oSheet2 = oWB.Worksheets.Add();
                        oSheet2.Name = "Summary";
                        oSheet2.Select();
                        oXL.ActiveSheet.Range["A1:B4"].Resize[transposedRange.GetUpperBound(0), transposedRange.GetUpperBound(1)] = transposedRange;

                        oXL.Visible = true;
                        oXL.UserControl = true;


                    }
                    catch (Exception theException)
                    {
                        string errorMessage;
                        errorMessage = "Error: ";
                        errorMessage = string.Concat(errorMessage, theException.Message);
                        errorMessage = string.Concat(errorMessage, " Line: ");
                        errorMessage = string.Concat(errorMessage, theException.Source);

                        MessageBox.Show(errorMessage, "Error");
                    }
                    
                }
            }
        }


        /// <summary>
        /// Launch Excel/Open if Launched.  Clear Sheet1 before import.
        /// </summary>
        /// <param name="oXL">Excel instance</param>
        /// <param name="oWB">Workbook Name</param> Hard Coded to vallen.xlsx, relative location
        /// <param name="oSheet">Worksheet Name</param>
        private static void ExcelLauncher(out Excel.Application oXL, out Excel._Workbook oWB, out Excel._Worksheet oSheet)
        {
            string filenameS = "vallen.xlsx";
            bool AppisOpened = true;
            bool WBisOpened = true;
            //test if App is open, else open it.
            string AppwasOpen = XLAppIsOpen().ToString();
            //test if WB is open. else, do nothing
            string WBwasOpen = WbIsOpened(filenameS).ToString();
            //if WB isn't open, try opening.  Else, try creating.
            if (!WbIsOpened(filenameS))
            {
                //oXL = new Excel.Application();  // Pretty sure this works
                try
                {
                    oXL = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    MessageBox.Show("Opening workbook.", "XLS is Open, Get WB"); //oXL.ActiveWorkbook.Name
                    oWB = (oXL.Workbooks.Open(filenameS));
                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                    oSheet.Cells.Clear();
                }
                catch (COMException ex)
                {
                    oXL = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    oXL.Visible = true;

                    if ((oXL.ActiveWorkbook ?? null) != null)
                        MessageBox.Show("Excel started. Active workbook being created: " + oXL.ActiveWorkbook.Name, " ...");
                    else
                        MessageBox.Show("Excel started. There is no open workbook, creating one now...");

                    oWB = oXL.Workbooks.Add(Missing.Value);
                    oWB.SaveAs(filenameS, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                    MessageBox.Show("COM Exception caught: " + ex.Message.ToString());
                }
            }
            //if WB is open, try setting our variables and clearing the active sheet, else try creating new WB (should never occur);
            else
            {
                try
                {
                    oXL = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    oWB = (oXL.Workbooks.get_Item(filenameS));
                    MessageBox.Show("Excel was running. Active workbook is:" + oXL.ActiveWorkbook.Name, "Already Running");
                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                    oSheet.Cells.Clear();
                }
                catch (COMException ex)
                {
                    oXL = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

                    oXL.Visible = true;
                    oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                    oWB.SaveAs(filenameS, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                    MessageBox.Show("Major Code Failure. Continuing. " + oXL.ActiveWorkbook.Name, "Started Excel.");
                    MessageBox.Show(ex.Message.ToString());
                }
            }



            
            bool WbIsOpened(string wbook)
            {

                Excel.Application exApp;
                exApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                try
                {
                    exApp.Workbooks.get_Item(wbook);
                }
                catch (Exception)
                {
                    WBisOpened = false;
                }

                return WBisOpened;
            }

            bool XLAppIsOpen()
            {
                Excel._Application xlObj;
                //Excel._Workbook oWBinternal;
                //Excel._Worksheet oSheetinternal;
                try
                {

                    xlObj = (Excel._Application)Marshal.GetActiveObject("Excel.Application");

                }
                catch (COMException ex)
                {
                    //If there is no running instance, it creates a new one
                    //Type type = Type.GetTypeFromProgID("Word.Application");
                    //word = System.Activator.CreateInstance(type);

                    xlObj = new Excel.Application();

                }
                return AppisOpened;
            }
            //object[] ExcelFileName = new object[1];
            //ExcelFileName[0] = new { Filename = "vallen.xlsx" };


        }




        /// <summary>
        /// Charting and Graphing
        /// </summary>
        /// <param name="oWS"></param>

        private void DisplayQuarterlySales(Excel._Worksheet oWS)
        {

            Excel._Workbook oWB;
            Excel.Series oSeries;
            Excel.Range oResizeRange;
            Excel._Chart oChart;
            string sMsg;
            int iNumQtrs;



            //Determine how many quarters to display data for.
            for (iNumQtrs = 4; iNumQtrs >= 2; iNumQtrs--)
            {
                sMsg = "Enter sales data for ";
                sMsg = string.Concat(sMsg, iNumQtrs);
                sMsg = string.Concat(sMsg, " quarter(s)?");

                DialogResult iRet = MessageBox.Show(sMsg, "Quarterly Sales?",
                MessageBoxButtons.YesNo);
                if (iRet == DialogResult.Yes)
                    break;
            }

            sMsg = "Displaying data for ";
            sMsg = string.Concat(sMsg, iNumQtrs);
            sMsg = string.Concat(sMsg, " quarter(s).");

            MessageBox.Show(sMsg, "Quarterly Sales");

            //Starting at E1, fill headers for the number of columns selected.
            oResizeRange = oWS.get_Range("E1", "E1").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=\"Q\" & COLUMN()-4 & CHAR(10) & \"Sales\"";

            //Change the Orientation and WrapText properties for the headers.
            oResizeRange.Orientation = 38;
            oResizeRange.WrapText = true;

            //Fill the interior color of the headers.
            oResizeRange.Interior.ColorIndex = 36;

            //Fill the columns with a formula and apply a number format.
            oResizeRange = oWS.get_Range("E2", "E6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=RAND()*100";
            oResizeRange.NumberFormat = "$0.00";

            //Apply borders to the Sales data and headers.
            oResizeRange = oWS.get_Range("E1", "E6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            //Add a Totals formula for the sales data and apply a border.
            oResizeRange = oWS.get_Range("E8", "E8").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=SUM(E2:E6)";
            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle
            = Excel.XlLineStyle.xlDouble;
            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight
            = Excel.XlBorderWeight.xlThick;

            //Add a Chart for the selected data.
            oWB = (Excel._Workbook)oWS.Parent;
            oChart = (Excel._Chart)oWB.Charts.Add(Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);

            //Use the ChartWizard to create a new chart from the selected data.
            oResizeRange = oWS.get_Range("A15:E489", Missing.Value).get_Resize(
            Missing.Value, iNumQtrs);
            oChart.ChartWizard(oResizeRange, Excel.XlChartType.xl3DColumn, Missing.Value,
            Excel.XlRowCol.xlColumns, Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            oSeries = (Excel.Series)oChart.SeriesCollection(1);
            oSeries.XValues = oWS.get_Range("A2", "A6");
            for (int iRet = 1; iRet <= iNumQtrs; iRet++)
            {
                oSeries = (Excel.Series)oChart.SeriesCollection(iRet);
                string seriesName;
                seriesName = "=\"Q";
                seriesName = string.Concat(seriesName, iRet);
                seriesName = string.Concat(seriesName, "\"");
                oSeries.Name = seriesName;
            }

            oChart.Location(Excel.XlChartLocation.xlLocationAsObject, oWS.Name);

            //Move the chart so as not to cover your data.
            oResizeRange = (Excel.Range)oWS.Rows.get_Item(10, Missing.Value);
            oWS.Shapes.Item("Chart 1").Top = (float)(double)oResizeRange.Top;
            oResizeRange = (Excel.Range)oWS.Columns.get_Item(2, Missing.Value);
            oWS.Shapes.Item("Chart 1").Left = (float)(double)oResizeRange.Left;
        }

        #endregion


        /// <summary>
        /// Helper to convert column number to column letter for Excel Integration
        /// </summary>
        const int ColumnBase = 26;
        const int DigitMax = 7; // ceil(log26(Int32.Max))
        const string Digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public static string IndexToColumn(int index)
        {
            if (index <= 0)
                throw new IndexOutOfRangeException("index must be a positive number");

            if (index <= ColumnBase)
                return Digits[index - 1].ToString();

            StringBuilder sb = new StringBuilder().Append(' ', DigitMax);
            int current = index;
            int offset = DigitMax;
            while (current > 0)
            {
                sb[--offset] = Digits[--current % ColumnBase];
                current /= ColumnBase;
            }
            return sb.ToString(offset, DigitMax - offset);
        }







    }
}
