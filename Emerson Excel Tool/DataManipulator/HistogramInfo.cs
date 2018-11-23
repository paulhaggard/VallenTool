using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelToolkit;
using ExcelToolkit.Statistics;
using static System.Windows.Forms.CheckedListBox;

namespace Emerson_Excel_Tool
{
    public partial class HistogramInfo : Form
    {
        #region Events

        public delegate void CalculationHandler(object sender, Histogram results);
        /// <summary>
        /// Triggered when the histogramInfo has completed
        /// </summary>
        public event CalculationHandler CompletionEvent;
        public void OnCompletionEvent()
        {
            CompletionEvent?.Invoke(this, histogram);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Dataset to create the histogram from
        /// </summary>
        private Dataset dataset { get; set; } = new Dataset();

        public Histogram histogram { get; }

        #endregion

        #region Constructors

        /// <summary>
        /// Creates an empty Histogram wizard
        /// ONLY USED FOR THE DEFAULT WINDOWS FORM CREATOR
        /// DO NOT USE!!!!
        /// </summary>
        public HistogramInfo()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Creates a new histogram with the dataset given
        /// </summary>
        /// <param name="data">Dataset to create the histogram with</param>
        public HistogramInfo(Dataset data)
        {
            InitializeComponent();

            dataset = data;
            histogram = new Histogram(data);
            setup();
        }

        #endregion

        #region Methods

        /// <summary>
        /// Sets up the controls and stuff for the histogram
        /// max/mins for the numerics
        /// Sets up the selection boxes for the radio buttons
        /// </summary>
        private void setup()
        {
            numericUpDownBinCreator.Maximum = (decimal)dataset.Responses.Max();
            numericUpDownBinCreator.Minimum = 0;
            numericUpDownBinCount.Maximum = 2147483647;
            numericUpDownBinCount.Minimum = 1;
            radioButtonDefault.Checked = true;
            radioButtonRefresh();
            listBoxRefresh();
        }

        /// <summary>
        /// Refreshes the radiobuttons
        /// </summary>
        private void radioButtonRefresh()
        {
            if(radioButtonDefault.Checked)
            {
                numericUpDownBinCreator.Enabled = false;
                buttonAdd.Enabled = false;
                numericUpDownBinCount.Enabled = true;
            }
            else
            {
                numericUpDownBinCount.Enabled = false;
                numericUpDownBinCreator.Enabled = true;
                buttonAdd.Enabled = true;
            }
        }

        private void listBoxRefresh()
        {
            listBoxBins.Items.Clear();
            foreach (double d in histogram.Bins)
                listBoxBins.Items.Add(d);
        }

        #endregion

        private void buttonCencel_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void radioButtonDefault_CheckedChanged(object sender, EventArgs e)
        {
            radioButtonRefresh();
        }

        private void buttonRemove_Click(object sender, EventArgs e)
        {
            if (!radioButtonCustom.Checked)
                radioButtonCustom.Checked = true;
            listBoxBins.Items.RemoveAt(listBoxBins.SelectedIndex);
            histogram.Bins = listBoxBins.Items.Cast<double>().ToList();
            listBoxRefresh();
        }

        private void numericUpDownBinCount_ValueChanged(object sender, EventArgs e)
        {
            histogram.BinCount = (int)numericUpDownBinCount.Value;
            listBoxRefresh();
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            listBoxBins.Items.Add(numericUpDownBinCreator.Value);
            histogram.Bins = listBoxBins.Items.Cast<double>().ToList();
            listBoxRefresh();
        }

        private void buttonAccept_Click(object sender, EventArgs e)
        {
            OnCompletionEvent();
            Dispose();
        }
    }
}
