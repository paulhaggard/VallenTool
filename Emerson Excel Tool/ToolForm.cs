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
        ///
        ///

        public int fileCount { get; set; }

        /// <summary>
        /// Count and identify which file paths have been chosen.
        /// </summary>
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
            MessageBox.Show(Convert.ToString(selectedFileList.Length), "Selected file(s) count:");
            fileCount = testFileList.Count;
        }
        
        #endregion

       

        #region Unused XML example


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
            SetFileToProcess();
            RunExcelProcess();
            EmptyTheFileList();
        }

        private void aboutBtn_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This is a file processing tool built to provide fast, comparative testing of Vallen acoustic emission sensor tests.  " +
                "It works with .txt files (produced by the Vallen test equipment) and processes them in Excel.", "About the Emerson Tool");
        }
        #endregion


        /// <summary>
        /// Call a MessageBox on formless .cs by:
        /// CallMessageBox $varname;
        /// $varname.message = my message;
        /// </summary>
        public struct CallMessageBox
        {
            public string message;
            public CallMessageBox(string var)
            {
                message = var;
                MessageBox.Show(message);
            }
        }

        
    }
}

