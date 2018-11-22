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
using ExcelToolkit;
using ExcelToolkit.DataFormatting;

namespace Emerson_Excel_Tool
{
    public partial class ToolForm : Form
    {
        #region Properties

        /// <summary>
        /// Create an Excel Interop Object to be used for all excel interactions.
        /// </summary>
        private ExcelPort excelObject { get; set; } = new ExcelPort(false);

        /// <summary>
        /// Contains the data outside of the listbox
        /// </summary>
        private List<IExcelData> datasets { get; set; } = new List<IExcelData>();

        #endregion

        public ToolForm()
        {
            // Set a few initial forms configuration settings
            InitializeComponent();

            datasets.Add(new ExcelWorkbookFormatter());

            InitializeOpenFileDialog();

            prepareDataGridView();
        }
        
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

        #region Methods

        /// <summary>
        /// Count and identify which file paths have been chosen.
        /// </summary>
        public void SetFileToProcess()
        {
            datasets.AddRange(FileSelectionListBox.Items.Cast<Dataset>());  // Adds the data in the listbox to the processing data
        }

        /// <summary>
        /// Action to multi-select .txt files for processing. 
        /// To be displayed in list box.
        /// </summary>
        private void FileSelectionHelper()
        {
            //DialogResult dr = this.openFileDialog1.ShowDialog();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                // Read the files
                foreach (string file in openFileDialog1.FileNames)
                {
                    // Create a List Item.
                    try
                    {
                        Dataset d = Dataset.CreateDataTableFromFile("", file);
                        FileSelectionListBox.Items.Add(d);
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
            openFileDialog1.Filter = "Text (*.txt)|*.txt|All files (*.*)|*.*";
            // Allow the user to select multiple images.
            openFileDialog1.Multiselect = true;
            openFileDialog1.Title = ".txt File Browser";

        }

        /// <summary>
        /// To prevent the files list from being executed on twice without emptying list,
        /// this action clears the list of file locations after the Excel processing is requested.
        /// </summary>
        private void EmptyTheFileList()
        {
            // Resets the data list
            datasets.Clear();
            datasets.Add(new ExcelWorkbookFormatter());
        }

        /// <summary>
        /// Action to remove items from listbox, including multi-selection
        /// </summary>
        private void SelectAndRemoveListItems()
        {
            List<int> indexToRemove = new List<int>();
            foreach (int index in FileSelectionListBox.SelectedIndices)
            {
                indexToRemove.Add(index);
            }
            indexToRemove.Reverse();
            foreach (int index in indexToRemove)
            {
                FileSelectionListBox.Items.RemoveAt(index);
            }
        }

        /// <summary>
        /// Prepares the datagridView object for use with the datasets
        /// </summary>
        private void prepareDataGridView()
        {
            dataGridView1.ColumnCount = 2;
            dataGridView1.Name = "Preview box";
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.Columns[0].Name = "Frequency";
            dataGridView1.Columns[1].Name = "Response";
            dataGridView1.Rows.Add(new string[2] { "Select a file to preview it", "" });
        }

        #endregion

        #region Events

        /// <summary>
        /// Reads in the currently selected list item and puts it into the preview pane
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FilesSelected_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            Dataset d = (Dataset)FileSelectionListBox.SelectedItem;
            string[,] dt = d.GetStringData();

            dataGridView1.Columns[0].Name = dt[0, 0];
            dataGridView1.Columns[1].Name = dt[0, 1];

            for(int i = 1; i < dt.GetLength(0); i++)
                dataGridView1.Rows.Add(new string[2] { dt[i, 0], dt[i, 1] });
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
            for (int i = 0; i < FileSelectionListBox.Items.Count; i++)
            {
                MessageBox.Show(FileSelectionListBox.Items[i].ToString(), "Test File Location is set to:");
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
        private void runExcelBtn(object sender, EventArgs e)
        {
            SetFileToProcess();

            // Opens excel if it's not already open
            if (!excelObject.isAppOpen)
                excelObject.OpenApp();

            // Opens the workbook if it hasn't been opened yet
            if (!excelObject.isWbOpen)
                excelObject.OpenWorkbook("vallen.xlsx");

            excelObject.setVisible(true);

            excelObject.writeData(datasets);
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
            if(excelObject.isAppOpen)
                excelObject.CloseApp();
            if (!IsDisposed)
                Dispose();  // Gets rid of this object instance
        }

        #endregion

    }
}

