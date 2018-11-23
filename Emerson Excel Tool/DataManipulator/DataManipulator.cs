using ExcelToolkit;
using ExcelToolkit.DataFormatting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Emerson_Excel_Tool
{
    public partial class DataManipulator : Form
    {
        #region Properties

        /// <summary>
        /// Dataset to have the operations performed on
        /// </summary>
        private Dataset dataset { get; set; } = new Dataset();

        #endregion

        #region Constructors

        /// <summary>
        /// Creates an empty DataManipulator
        /// ONLY USED FOR THE DEFAULT WINDOWS FORM CREATOR
        /// DO NOT USE!!!!
        /// </summary>
        public DataManipulator()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Creates a DataManipulator using the data in the dataset supplied
        /// </summary>
        /// <param name="data">The data for the operations to be performed on</param>
        public DataManipulator(Dataset data)
        {
            InitializeComponent();

            dataset = data;
            refreshDataGrid();
        }

        #endregion

        #region Methods

        /// <summary>
        /// Refreshed the DataGridView to contain the dataset passed in when the window was created
        /// </summary>
        private void refreshDataGrid()
        {
            dataGridView1.ColumnCount = 2;
            dataGridView1.Rows.Clear();
            string[,] dt = dataset.GetStringData();

            dataGridView1.Columns[0].Name = "Frequency";
            dataGridView1.Columns[1].Name = "Response";

            for (int i = 14; i < dt.GetLength(0); i++)
                dataGridView1.Rows.Add(new string[2] { dt[i, 0], dt[i, 1] });
        }

        /// <summary>
        /// Plots the data provided onto a chart
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        private void MrPlotter<T>(IDataManData<T> data)
        {

        }

        #endregion

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        #region Histogram

        private void histogramToolStripMenuItem_Click(object sender, EventArgs e)
        {
            HistogramInfo histoWindow = new HistogramInfo(dataset);
            histoWindow.CompletionEvent += HistoWindow_CompletionEvent;
            histoWindow.Visible = true;
        }

        private void HistoWindow_CompletionEvent(object sender, IDataManData<double> results)
        {
            //TODO plot the histogram
        }

        #endregion
    }
}
