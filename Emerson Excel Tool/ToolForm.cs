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
        /// <summary>
        /// ToolForm constructor.  Sets a few defaults for the class.
        /// </summary>
        public ToolForm()
        {
            // Set a few initial form configuration settings
            InitializeComponent();
            InitializeOpenFileDialog();
        }


        // Create lists to store our files for Excel processing
        public static List<FileStats> appFileList = new List<FileStats>();
        public static List<DataSet> listOfDataSets = new List<DataSet>();
        public static List<DataTable> listOfDataTables = new List<DataTable>();


        // Form is loaded here.
        private void Form1_Load(object sender, EventArgs e)
        {
            // When the form loads, launch Excel.
            FileSelectionListBox.DisplayMember = "FileFullPath";
            Launch();
        }

        /// <summary>
        /// Method to get Excel running.
        /// </summary>FilesSelected datasource
        /// 
        public void Launch()
        {
            //launch excel, set local variables for app, workbook, sheet
            ExcelLauncher(out oXL, out oWB, out oSheet);
        }


        #region Unused Form Objects/Buttons for Events

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void folderBrowserDialog1_HelpRequest_1(object sender, EventArgs e)
        {

        }

        private void FilesSelected_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void helloWorldLabel_Click(object sender, EventArgs e)
        {

        }

        private void bindingNavigator1_RefreshItems(object sender, EventArgs e)
        {

        }

        #endregion


        #region FileSelectionSet
        ///
        ///

        public static int selectedFilesCount { get; set; }

        /// <summary>
        /// Count and identify which file paths have been chosen.
        /// </summary>
        public void SetFileToProcess()
        { selectedFilesCount = FileSelectionListBox.Items.Count; }
        /*    //Select all imported files.  Yes, this is dumb.
            FileSelectionListBox.Visible = false;
            for (int i = 0; i < FileSelectionListBox.Items.Count; i++)
            {
                FileSelectionListBox.SetSelected(i, true);
            }
            FileSelectionListBox.Visible = true;
            //Create array and fill with the strings of each file location
            String[] selectedFilesList = new string[FileSelectionListBox.Items.Count];
            FileSelectionListBox.SelectedValue CopyTo(selectedFilesList, 0);
            //Add each line of this array to a list.  Why?  Why not.
            for (int i = 0; i < selectedFilesList.Length; i++)
            {

                if (!testFileList.Any(e => e.Equals(selectedFilesList[i])))  //add only if DNE
                    if (!String.IsNullOrEmpty(selectedFilesList[i]))
                    {
                        {
                            testFileList.Add(selectedFilesList[i]);
                        }
                    }

            }
            //MessageBox.Show(Convert.ToString(selectedFilesList.Length), "Selected file(s) count:");
            selectedFilesCount = testFileList.Count;
        }*/

        #endregion



        #region Unused XML example





        #endregion

        #region Reading .txt line by line


        /* 
         * public void ReadTextFile()
         {
             string line;
             try
             {
                 this.InputText.Clear();
                 //Pass the file path and file name to the StreamReader constructor
                 StreamReader sr = new StreamReader("C:\\temp\\Jamaica.txt");

                 //Read the first line of text
                 line = sr.ReadLine();

                 //Continue to read until you reach end of file
                 while (line != null)
                 {
                     //write the lie to console window
                     //this.InputText.SelectionStart = InputText.Text.Length;

                     this.InputText.AppendText(line);
                     this.InputText.AppendText(Environment.NewLine);

                     //Read the next line
                     line = sr.ReadLine();
                 }

                 //close the file
                 sr.Close();
                 String dt;
                 DataSet_Processing instance = new DataSet_Processing();
                 instance.CreateDataTableFromFile(testFileLocation, testFileLocation);  ///create the table, this is expecting a table name, too
             }
             catch (Exception e)
             {
                 MessageBox.Show("Exception: " + e.Message);
             }
             finally
             {
                 MessageBox.Show("Executing finally block.");
             }
         }
         */


        /// <summary>
        /// File selector
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }





        #endregion

        #region Active Buttons

        private void testbuttn_Click(object sender, EventArgs e)
        {
            LoadDatagrid(out DataTable dt);

            // GetTableData(out listOfDataTables.ElementAt(0));
            dataGridViewer.AutoGenerateColumns = true;
            dataGridViewer.DataSource = dt;
        }

        private void testbuttn2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < appFileList.Count; i++)
            {

                MessageBox.Show(appFileList[i].FileFullPath.ToString(), "Test File Location is set to:");
            }
        }
        private void openFilesButton_Click(object sender, EventArgs e)
        {
            FileSelectionHelper();
        }


        private void RemoveFilesSelected_Click(object sender, EventArgs e)
        {
            SelectAndRemoveListItems();

        }


        private void runExcelBtn(object sender, System.EventArgs e)
        {
            SetFileToProcess();
            Launch();
            ProcessFilesinExcel();
            EmptyTheFileList();
        }

        private void aboutBtn_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This is a file processing tool built to provide fast, comparative testing of Vallen acoustic emission sensor test data.\n\n" +
                "It works with .txt files (exported by the Vallen sensor testing software) and processes them in Excel for statistical comparisons.\r\rVersion 1.10.25        ©2018", "About the Emerson Tool");
        }

        private void ToolForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Close();
        }



        private void dataGridViewer_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        #endregion

        #region deClassified Code

        #region     DataSet Processing


        //private string _tableName;
        //private string _tableFileLocation;
        DataSet _vallenLogs = new DataSet("Vallen Logs");

        /// <summary>
        /// CTOR: Create new dataset for our logs.  If set exists, do not create.
        /// </summary>

        #region Data Set & Data Tables Code

        ///Create a DataSet
        ///
        public static DataSet datasetName  // this is a property with an accessor.
        {
            get
            {
                return datasetName;
            }
            private set
            {
                datasetName = value;
            }
        }

        //set the working table file name and file path during processing of each file
        public string tableName { get; set; }
        public string tableFileLocation { get; set; }


        /// <summary>
        /// Write a datatable
        /// </summary>
        /// <param name="tableName">name of table</param>
        /// <param name="tableFileLocation">location of table on disk</param>
        /// <param name="dt"></param>
        /// <param name="results"></param>
        public void GetTableData(string tableName, string tableFileLocation, out DataTable dt, out string[,] results)
        {

            /////GOT HUNGUP HERE BADLY! READ MORE.

            dt = CreateDataTableFromFile(tableName, tableFileLocation);  //TODO come back and use name field
            results = new string[dt.Rows.Count, dt.Columns.Count];
            for (int index = 0; index < dt.Rows.Count; index++)
            {
                if (index == 0) { results[0, 0] = tableName; }
                else
                {
                    for (int columnIndex = 0; columnIndex < dt.Columns.Count; columnIndex++)
                    {
                        results[index, columnIndex] = dt.Rows[index][columnIndex].ToString();
                    }
                }
            }
        }
        public void GetTableData(out DataTable dt, out string[,] results)
        {

            /////GOT HUNGUP HERE BADLY! READ MORE.

            dt = CreateDataTableFromFile(tableName, tableFileLocation);  //TODO come back and use name field
            results = new string[dt.Rows.Count, dt.Columns.Count];
            for (int index = 0; index < dt.Rows.Count; index++)
            {
                if (index == 0) { results[0, 0] = tableName; }
                else
                {
                    for (int columnIndex = 0; columnIndex < dt.Columns.Count; columnIndex++)
                    {
                        results[index, columnIndex] = dt.Rows[index][columnIndex].ToString();
                    }
                }
            }
        }
        public void GetTableData(out DataTable dt)
        {

            /////GOT HUNGUP HERE BADLY! READ MORE.

            dt = this.CreateDataTableFromFile(tableName, tableFileLocation);  //TODO come back and use name field
            string[,] results = new string[dt.Rows.Count, dt.Columns.Count];
            for (int index = 0; index < dt.Rows.Count; index++)
            {
                if (index == 0) { results[0, 0] = tableName; }
                else
                {
                    for (int columnIndex = 0; columnIndex < dt.Columns.Count; columnIndex++)
                    {
                        results[index, columnIndex] = dt.Rows[index][columnIndex].ToString();
                    }
                }
            }
        }




        /// <summary>
        /// Create a data table from a file.
        /// </summary>
        /// <returns></returns>


        public DataTable CreateDataTableFromFile(string _name, string fileLocation)
        {


            DataTable newTable = new DataTable();
            newTable = _vallenLogs.Tables.Add(_name);

            DataColumn FreqColumn =
            newTable.Columns.Add("Frequency", typeof(string));
            newTable.Columns.Add("Response", typeof(string));
            newTable.PrimaryKey = new DataColumn[] { FreqColumn };
            DataRow dr;


            StreamReader sr = new StreamReader(fileLocation);
            string input;
            while ((input = sr.ReadLine()) != null)
            {
                if (input == string.Empty)
                {
                    continue;
                }
                else
                {

                    string[] s = input.Split(new char[] { '\t' });
                    dr = newTable.NewRow();
                    dr["Frequency"] = s[0];
                    dr["Response"] = s[1];
                    newTable.Rows.Add(dr);
                }
            }
            sr.Close();
            return newTable;
        }


        /// <summary>
        /// DataTable for List
        /// </summary>
        /// <param name="fileLocation"></param>
        /// <returns></returns>
        public DataTable CreateDataTableFromFile(List<string> fileLocation)
        {



            DataTable newtable = new DataTable();
            DataColumn dc;
            DataRow dr;
            for (int i = 0; i < fileLocation.Count; i++)
            {
                dc = new DataColumn();
                dc.DataType = System.Type.GetType("System.String");
                dc.ColumnName = "c1";
                dc.Unique = false;
                newtable.Columns.Add(dc);
                dc = new DataColumn();
                dc.DataType = System.Type.GetType("System.String");
                dc.ColumnName = "c2";
                dc.Unique = false;
                newtable.Columns.Add(dc);

                StreamReader sr = new StreamReader(fileLocation[i]);
                string input;
                while ((input = sr.ReadLine()) != null)
                {
                    if (input == string.Empty)
                    {
                        continue;
                    }
                    else
                    {

                        string[] s = input.Split(new char[] { '\t' });
                        dr = newtable.NewRow();
                        dr["c1"] = s[0];
                        dr["c2"] = s[1];
                        newtable.Rows.Add(dr);
                    }
                }
                sr.Close();
            }
            return newtable;
        }


        /// <summary>
        /// Load selected data into Datagrid from DataTables
        /// </summary>
        /// <param name="dt"></param>
        public void LoadDatagrid(out DataTable dt)
        {
            MessageBox.Show(listOfDataTables.Count.ToString());
            MessageBox.Show(listOfDataTables.ElementAt(0).ToString());
            MessageBox.Show(listOfDataTables.ElementAt(0).TableName);
            DataTable dt_int = new DataTable();
            string[,] results = new string[dt_int.Rows.Count, dt_int.Columns.Count];
            //listOfDataTables.ElementAt(0).GetTableData(out dt_int, out results);
            dt = dt_int;

        }

        #endregion

        #endregion



        #region Excel Link

        //"http://aka.ms/dotnet-get-started-desktop");
        // initialize objects for Excel
        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;
        Excel.Range oRng;


        /// <summary>
        /// Launch Excel/Open if Launched. Open/create Workbook. Clear Sheet1 before import.
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
                try
                {
                    oXL = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    //MessageBox.Show("Opening workbook.", "XLS is Open, Get WB"); //oXL.ActiveWorkbook.Name
                    oWB = (oXL.Workbooks.Open(filenameS));
                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                    oSheet.Cells.Clear();
                }
                catch (COMException ex)
                {
                    oXL = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    oXL.Visible = true;
                    MessageBox.Show("Excel started. Active workbook being created: " + oXL.ActiveWorkbook.Name, " ...");
                    oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
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
                    oXL = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    oWB = (oXL.Workbooks.get_Item(filenameS));
                    //MessageBox.Show("Excel was running. Active workbook is:" + oXL.ActiveWorkbook.Name, "Already Running");
                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                    oSheet.Cells.Clear();
                }
                catch (COMException ex)
                {
                    oXL = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                    oXL.Visible = true;
                    oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                    oWB.SaveAs(filenameS, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                    MessageBox.Show("Major Code Failure. Continuing. " + oXL.ActiveWorkbook.Name, "Started Excel.");
                    MessageBox.Show(ex.Message.ToString());
                }
            }



            // Test if workbook is already open
            bool WbIsOpened(string wbook)
            {

                Excel.Application exApp;
                exApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
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

            // Test if Excel is already open
            bool XLAppIsOpen()
            {
                Excel._Application xlObj;
                Excel._Workbook oWBinternal;
                Excel._Worksheet oSheetinternal;
                try
                {

                    xlObj = (Excel._Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

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



        }


        /// <summary>
        /// Close Excel file without save, end Interop COM.
        /// </summary>
        public void Close()
        {
            try
            {
                if (oXL != null)
                {
                    oXL.Visible = true;
                    bool SaveChanges;
                    oWB.Close(SaveChanges = false, Type.Missing, Type.Missing);
                    oXL.Quit();
                    oXL = null;
                }

            }
            catch (Exception ex)
            {
                // alternatively you can show Error message
                MessageBox.Show("Excel Sheet has already been closed. Exiting tool...");
            }
            finally
            {
                // release ref vars
                if (oXL != null)
                    try
                    {
                        oXL.Visible = true;
                        oXL.Quit();
                        oSheet = null;
                        oWB = null;
                        oXL = null;

                    }
                    catch (Exception ex)
                    { MessageBox.Show("Alert: Excel crashed."); }

            }
        }

        /// <summary>
        /// Process Imported txt Data, Run Excel functions
        /// </summary>
        public void ProcessFilesinExcel()
        {
            if (selectedFilesCount < 1)
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



                    int columnNumb = 2 * selectedFilesCount;
                    for (int i = 0; i < selectedFilesCount; i++)
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



                    for (int i = 0; i < selectedFilesCount; i++)
                    {
                        listOfDataTables.Add(new DataTable { TableName = "File #" + (i + 1) });
                    }
                    int totalDataTablesLoaded = listOfDataTables.Count;
                    int _filecounter = 0;
                    for
                        (int i = 0; i < (2 * listOfDataTables.Count); i = i + 2)
                    {

                        listOfDataTables.ElementAt(_filecounter).TableName = Path.GetFileName(appFileList.ElementAt(_filecounter).FileName.ToString());
                        listOfDataTables.ElementAt(_filecounter).Namespace = appFileList.ElementAt(_filecounter).FileFullPath.ToString();
                        GetTableData(listOfDataTables.ElementAt(_filecounter).TableName, listOfDataTables.ElementAt(_filecounter).Namespace, out dt, out results);
                        oSheet.get_Range(IndexToColumn(i + 1) + dt.Columns.Count, IndexToColumn(i + 2) + dt.Rows.Count).Value2 = results;
                        _filecounter++;

                    }

                    oRng = oSheet.Range["A1:D489"];
                    Object[,] transposedRange = (Object[,])oXL.WorksheetFunction.Transpose(oRng.Value2);
                    Excel.Worksheet oSheet2;
                    oSheet2 = oWB.Worksheets.Add();
                    oSheet2.Name = "Sheet2";
                    oSheet2.Select();
                    oXL.ActiveSheet.Range["A1:B4"].Resize[transposedRange.GetUpperBound(0), transposedRange.GetUpperBound(1)] = transposedRange;

                    oXL.Visible = true;
                    oXL.UserControl = true;


                }
                catch (Exception theException)
                {
                    String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);

                    MessageBox.Show(errorMessage, "Error");
                }
            }
        }








        /// <summary>
        /// Charting and Graphing example
        /// </summary>
        /// <param name="oWS"></param>
        private void DisplayQuarterlySales(Excel._Worksheet oWS)
        {

            Excel._Workbook oWB;
            Excel.Series oSeries;
            Excel.Range oResizeRange;
            Excel._Chart oChart;
            String sMsg;
            int iNumQtrs;



            //Determine how many quarters to display data for.
            for (iNumQtrs = 4; iNumQtrs >= 2; iNumQtrs--)
            {
                sMsg = "Enter sales data for ";
                sMsg = String.Concat(sMsg, iNumQtrs);
                sMsg = String.Concat(sMsg, " quarter(s)?");

                DialogResult iRet = MessageBox.Show(sMsg, "Quarterly Sales?",
                MessageBoxButtons.YesNo);
                if (iRet == DialogResult.Yes)
                    break;
            }

            sMsg = "Displaying data for ";
            sMsg = String.Concat(sMsg, iNumQtrs);
            sMsg = String.Concat(sMsg, " quarter(s).");

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
                String seriesName;
                seriesName = "=\"Q";
                seriesName = String.Concat(seriesName, iRet);
                seriesName = String.Concat(seriesName, "\"");
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

        #region Column number to Letter Excel Tool
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

            var sb = new StringBuilder().Append(' ', DigitMax);
            var current = index;
            var offset = DigitMax;
            while (current > 0)
            {
                sb[--offset] = Digits[--current % ColumnBase];
                current /= ColumnBase;
            }
            return sb.ToString(offset, DigitMax - offset);
        }
        #endregion






        #endregion

    }
}

