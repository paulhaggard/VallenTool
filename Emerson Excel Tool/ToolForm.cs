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
    public partial class ToolForm : Form
    {
        public ToolForm()
        {
            // Set a few initial forms configuration settings
            InitializeComponent();
            InitializeOpenFileDialog();       
        }


        // Create an Excel Interop Object to be used for all excel interactions.
        ExcelProcess excelObject = new ExcelProcess();

        // Create a list to store our files for Excel processing
        public List<string> testFileList = new List<string>();
        
        #region Unused Form Objects/Buttons for Events
        
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
        ///
        ///

        public int selectedFilesCount { get; set; }

        /// <summary>
        /// Count and identify which file paths have been chosen.
        /// </summary>
        public void SetFileToProcess()
        {
            /*
            //Select all imported files.  Yes, this is dumb.
            FileSelectionListBox.Visible = false;
            for (int i = 0; i < FileSelectionListBox.Items.Count; i++)
            {
                FileSelectionListBox.SetSelected(i, true);
            }
            FileSelectionListBox.Visible = true;
            //Create array and fill with the strings of each file location
            String[] selectedFilesList = new string[FileSelectionListBox.Items.Count];
            FileSelectionListBox.SelectedItems.CopyTo(selectedFilesList, 0);
            */

            // I'm pretty sure this is equivalent to what you wrote

            List<string> selectedFilesList = new List<string>(FileSelectionListBox.Items.Count);

            foreach (string s in FileSelectionListBox.Items)
                selectedFilesList.Add(s);

            //Add each line of this array to a list.  Why?  Why not.
            foreach(string s in selectedFilesList)
            {

                if (!testFileList.Any(e => e.Equals(s)))  //add only if DNE
                    if (!string.IsNullOrEmpty(s))
                    { 
                        testFileList.Add(s);
                    }

            }
            //MessageBox.Show(Convert.ToString(selectedFilesList.Length), "Selected file(s) count:");
            selectedFilesCount = testFileList.Count;
        }
        
        #endregion

       

        #region Unused XML example


        void StoreFilesList(FilesList filesList)
        {
            // var is for lazy people, use the actual type definition
            XmlDocument doc = new XmlDocument();
            doc.Load(filesList.FileName);

            XmlElement channel = doc["rss"]["channel"];
            XmlNodeList items = channel.GetElementsByTagName("item");
            filesList.FileLocation = channel["title"].InnerText;
            filesList.Link = channel["link"].InnerText;
            filesList.Date = channel["description"].InnerText;

        }


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

        /// <summary>
        /// Triggered when the test button 1 is pressed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void testbuttn_Click(object sender, EventArgs e)
        {
            SetFileToProcess();

        }

        /// <summary>
        /// Triggered when the test button 2 is pressed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void testbuttn2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < testFileList.Count; i++)
            {

                MessageBox.Show(testFileList[i], "Test File Location is set to:");
            }
        }

        /// <summary>
        /// Opens the file selection dialog
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openFilesButton_Click(object sender, EventArgs e)
        {
            FileSelectionHelper();
        }

        /// <summary>
        /// Clears the selected files from the selection list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RemoveFilesSelected_Click(object sender, EventArgs e)
        {
            SelectAndRemoveListItems();
        }

        /// <summary>
        /// Processes the selected files and inserts them into excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void runExcelBtn(object sender, System.EventArgs e)
        {
            SetFileToProcess();
            excelObject.Launch(testFileList);
            EmptyTheFileList();
        }

        /// <summary>
        /// Opens a dialog that displays information about the purpose of the software
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void aboutBtn_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This is a file processing tool built to provide fast, comparative testing of Vallen acoustic emission sensor tests.  " +
                "It works with .txt files (produced by the Vallen test equipment) and processes them in Excel.", "About the Emerson Tool");
        }

        /// <summary>
        /// Triggered when a the 'x' button is pressed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            excelObject.Close();
            if (!IsDisposed)
                Dispose();  // Gets rid of this object instance
        }

        #endregion

    }
}

