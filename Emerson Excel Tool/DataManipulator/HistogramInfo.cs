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

namespace Emerson_Excel_Tool
{
    public partial class HistogramInfo : Form
    {
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
        }

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
        }

        private void numericUpDownBinCount_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
