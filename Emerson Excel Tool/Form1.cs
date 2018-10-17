using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

namespace Emerson_Excel_Tool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            InitializeOpenFileDialog();
        }


        public List<string> testFileList = new List<string>();
        public string testFileLocation = @"C:\Projects\2007 - Emerson AE\08. Testing\Paul Report Writing\Vallen Sensor Tests\296 a-b\VS900-RIC - oil - vallen reset.r216.txt";

        #region Unused Form Objects for Events
        //is this Form1 Load 1 needed? review.
        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

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
        ///Select the first file in the list to process. sort of works.  Doesn't account for files that are removed after pressing 'process' if 'process' is run again.
        ///
        int fileCount = new int();
        public void SetFileToProcess()
        {
            //Select all imported files.  Yes, this is dumb.
            FilesSelected.Visible = false;
            for (int i = 0; i < FilesSelected.Items.Count; i++)
            {
                FilesSelected.SetSelected(i, true);
            }
            FilesSelected.Visible = true;
            //Create array and fill with the strings of each file location
            String[] selectedFileList = new string[FilesSelected.Items.Count];
            FilesSelected.SelectedItems.CopyTo(selectedFileList, 0);
            //Add each line of this array to a list.  Why?  Why not.
            for (int i = 0; i < selectedFileList.Length; i++)
            {

                if (!testFileList.Any(e => e.Equals(selectedFileList[i])))  //add only if DNE
                    if (!String.IsNullOrEmpty(selectedFileList[i]))
                    {
                        {
                            testFileList.Add(selectedFileList[i]);
                        }
                    }

            }
            MessageBox.Show(Convert.ToString(selectedFileList.Length), "Selected files in array are:");
            
        }
        
        #endregion

        #region Excel Linking

        /// <summary>
        /// excel linking methods are listed here.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // Click on the link below to continue learning how to build a desktop app using WinForms!
            System.Diagnostics.Process.Start("http://aka.ms/dotnet-get-started-desktop");

        }


        private void RunExcelProcess()
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "Test 1";
                oSheet.Cells[1, 2] = "Response 1";
                oSheet.Cells[1, 3] = "Full Name";
                oSheet.Cells[1, 4] = "Salary";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "D1").Font.Bold = true;
                oSheet.get_Range("A1", "D1").VerticalAlignment =
                Excel.XlVAlign.xlVAlignCenter;

                /*
                // Create an array to multiple values at once.
                string[,] saNames = new string[5, 2];

                saNames[0, 0] = "John";
                saNames[0, 1] = "Smith";
                saNames[1, 0] = "Tom";
                saNames[1, 1] = "Brown";
                saNames[2, 0] = "Sue";
                saNames[2, 1] = "Thomas";
                saNames[3, 0] = "Jane";
                saNames[3, 1] = "Jones";
                saNames[4, 0] = "Adam";
                saNames[4, 1] = "Johnson"; 
                */

                //Fill A2:B6 with an array of values (First and Last Names).
                //oSheet.get_Range("A2", "B6").Value2 = saNames;

                //attempt to fill excel with DT object.
                DataTable dt;
                string[,] results;
                DataSet_Processing dataset1 = new DataSet_Processing();
                dataset1.tableFileLocation = testFileList.ElementAt(0);
                dataset1.GetTableData(out dt, out results);
                

                oSheet.get_Range("A" + dt.Columns.Count, "B" + dt.Rows.Count).Value2 = results;

                DataSet_Processing dataset2 = new DataSet_Processing();
                dataset2.tableFileLocation = testFileList.ElementAt(1);
                dataset2.GetTableData(out dt, out results);

                oSheet.get_Range("C" + dt.Columns.Count, "D" + dt.Rows.Count).Value2 = results;
                /*
                //Fill C2:C6 with a relative formula (=A2 & " " & B2).
                oRng = oSheet.get_Range("C2", "C6");
                oRng.Formula = "=A2 & \" \" & B2";

                //Fill D2:D6 with a formula(=RAND()*100000) and apply format.
                oRng = oSheet.get_Range("D2", "D6");
                oRng.Formula = "=RAND()*100000";
                oRng.NumberFormat = "$0.00";

                //AutoFit columns A:D.
                oRng = oSheet.get_Range("A1", "D1");
                oRng.EntireColumn.AutoFit();

                //Manipulate a variable number of columns for Quarterly Sales Data.
                DisplayQuarterlySales(oSheet);
                */
                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
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
            oResizeRange = oWS.get_Range("E2:E6", Missing.Value).get_Resize(
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

        #region File Selection Methods


        void StoreFilesList(FilesList filesList)
        {
            var doc = new XmlDocument();
            doc.Load(filesList.FileName);

            XmlElement channel = doc["rss"]["channel"];
            XmlNodeList items = channel.GetElementsByTagName("item");
            filesList.FileLocation = channel["title"].InnerText;
            filesList.Link = channel["link"].InnerText;
            filesList.Date = channel["description"].InnerText;

        }


        #endregion

        #region Reading .txt line by line


        public void ReadTextFile()
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



        /// <summary>
        /// File selector
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }





        #endregion

        #region Buttons

        private void testbuttn_Click(object sender, EventArgs e)
        {
            SetFileToProcess();

        }

        private void testbuttn2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < testFileList.Count; i++)
            {

                MessageBox.Show(testFileList[i], "Test File Location is set to:");
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
            RunExcelProcess();
            EmptyTheFileList();
        }

        private void aboutBtn_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This is a file processing tool built to provide fast, comparative testing of Vallen acoustic emission sensor tests.  " +
                "It works with .txt files (produced by the Vallen test equipment) and processes them in Excel.", "About the Emerson Tool");
        }
        #endregion

        #region Actions

        /// <summary>
        /// Action to remove items from listbox, including multi-selection
        /// </summary>
        private void SelectAndRemoveListItems()
        {
            List<int> indexToRemove = new List<int>();
            foreach (int index in FilesSelected.SelectedIndices)
            {
                indexToRemove.Add(index);
            }
            indexToRemove.Reverse();
            foreach (int index in indexToRemove)
            {
                FilesSelected.Items.RemoveAt(index);
            }
        }

        /// <summary>
        /// Action to multi-select .txt files for processing. 
        /// To be displaye din list box.
        /// </summary>
        private void FileSelectionHelper()
        {
            //DialogResult dr = this.openFileDialog1.ShowDialog();
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                // Read the files
                foreach (String file in openFileDialog1.FileNames)
                {
                    // Create a List Item.
                    try
                    {
                        FilesSelected.Items.Add(file);
                    }

                    catch (Exception ex)
                    {
                        // Could not load the file - probably related to Windows file system permissions.
                        MessageBox.Show("Cannot display the image: " + file.Substring(file.LastIndexOf('\\'))
                            + ". You may not have permission to read the file, or " +
                            "it may be corrupt.\n\nReported error: " + ex.Message);
                    }
                }
            }
        }

        /// <summary>
        /// Sets the file Dialog settings appropriately for app.
        /// </summary>
        private void InitializeOpenFileDialog()
        {
            this.openFileDialog1.Filter =
        "Text (*.txt)|*.txt|All files (*.*)|*.*";
            // Allow the user to select multiple images.
            this.openFileDialog1.Multiselect = true;
            this.openFileDialog1.Title = ".txt File Browser";

        }

        /// <summary>
        /// To prevent the files list from being executed on twice without emptying list,
        /// this action clears the list of file locations after the Excel processing is requested.
        /// </summary>
        private void EmptyTheFileList()
        {
            testFileList.Clear();
        }

        #endregion


    }
}

