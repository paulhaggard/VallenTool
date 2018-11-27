namespace Emerson_Excel_Tool
{
    partial class ToolForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Designer
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.runExcelProcess = new System.Windows.Forms.Button();
            this.helloWorldLabel = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.FileSelectionListBox = new System.Windows.Forms.ListBox();
            this.InputText = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.openFilesButton = new System.Windows.Forms.Button();
            this.RemoveFilesSelected = new System.Windows.Forms.Button();
            this.aboutBtn = new System.Windows.Forms.Button();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.dataGridViewer = new System.Windows.Forms.DataGridView();
            this.testbuttn = new System.Windows.Forms.Button();
            this.testbuttn2 = new System.Windows.Forms.Button();
            this.tagsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.filesListBindingSource = new System.Windows.Forms.BindingSource(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewer)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tagsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.filesListBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // runExcelProcess
            // 
            this.runExcelProcess.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.runExcelProcess.Location = new System.Drawing.Point(790, 478);
            this.runExcelProcess.Margin = new System.Windows.Forms.Padding(2);
            this.runExcelProcess.Name = "runExcelProcess";
            this.runExcelProcess.Size = new System.Drawing.Size(152, 41);
            this.runExcelProcess.TabIndex = 2;
            this.runExcelProcess.Text = "Process Loaded Files to Excel";
            this.runExcelProcess.UseVisualStyleBackColor = true;
            this.runExcelProcess.Click += new System.EventHandler(this.runExcelBtn);
            // 
            // helloWorldLabel
            // 
            this.helloWorldLabel.AutoSize = true;
            this.helloWorldLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.helloWorldLabel.Location = new System.Drawing.Point(312, 16);
            this.helloWorldLabel.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.helloWorldLabel.Name = "helloWorldLabel";
            this.helloWorldLabel.Size = new System.Drawing.Size(203, 26);
            this.helloWorldLabel.TabIndex = 3;
            this.helloWorldLabel.Text = "Files for Processing";
            this.helloWorldLabel.Click += new System.EventHandler(this.helloWorldLabel_Click);
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.Description = "Select Folder to Import";
            this.folderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.MyDocuments;
            this.folderBrowserDialog1.ShowNewFolderButton = false;
            this.folderBrowserDialog1.HelpRequest += new System.EventHandler(this.folderBrowserDialog1_HelpRequest_1);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.AddExtension = false;
            this.openFileDialog1.DefaultExt = "txt";
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "*.txt|";
            this.openFileDialog1.Multiselect = true;
            this.openFileDialog1.SupportMultiDottedExtensions = true;
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // FileSelectionListBox
            // 
            this.FileSelectionListBox.AllowDrop = true;
            this.FileSelectionListBox.BackColor = System.Drawing.SystemColors.Window;
            this.FileSelectionListBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.FileSelectionListBox.FormattingEnabled = true;
            this.FileSelectionListBox.HorizontalScrollbar = true;
            this.FileSelectionListBox.Location = new System.Drawing.Point(0, 0);
            this.FileSelectionListBox.Name = "FileSelectionListBox";
            this.FileSelectionListBox.ScrollAlwaysVisible = true;
            this.FileSelectionListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.FileSelectionListBox.Size = new System.Drawing.Size(625, 424);
            this.FileSelectionListBox.Sorted = true;
            this.FileSelectionListBox.TabIndex = 5;
            this.FileSelectionListBox.SelectedIndexChanged += new System.EventHandler(this.FilesSelected_SelectedIndexChanged);
            // 
            // InputText
            // 
            this.InputText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.InputText.Font = new System.Drawing.Font("Microsoft Sans Serif", 25F);
            this.InputText.Location = new System.Drawing.Point(0, 0);
            this.InputText.Multiline = true;
            this.InputText.Name = "InputText";
            this.InputText.Size = new System.Drawing.Size(314, 424);
            this.InputText.TabIndex = 6;
            this.InputText.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(181, 20);
            this.label1.TabIndex = 7;
            this.label1.Text = "Text Preview Area (beta)";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // openFilesButton
            // 
            this.openFilesButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.openFilesButton.Location = new System.Drawing.Point(317, 479);
            this.openFilesButton.Name = "openFilesButton";
            this.openFilesButton.Size = new System.Drawing.Size(154, 39);
            this.openFilesButton.TabIndex = 11;
            this.openFilesButton.Text = "Select files for processing...";
            this.openFilesButton.UseVisualStyleBackColor = true;
            this.openFilesButton.Click += new System.EventHandler(this.openFilesButton_Click);
            // 
            // RemoveFilesSelected
            // 
            this.RemoveFilesSelected.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.RemoveFilesSelected.Location = new System.Drawing.Point(806, 22);
            this.RemoveFilesSelected.Name = "RemoveFilesSelected";
            this.RemoveFilesSelected.Size = new System.Drawing.Size(128, 23);
            this.RemoveFilesSelected.TabIndex = 12;
            this.RemoveFilesSelected.Text = "Remove Selected";
            this.RemoveFilesSelected.UseVisualStyleBackColor = true;
            this.RemoveFilesSelected.Click += new System.EventHandler(this.RemoveFilesSelected_Click);
            // 
            // aboutBtn
            // 
            this.aboutBtn.Location = new System.Drawing.Point(16, 481);
            this.aboutBtn.Name = "aboutBtn";
            this.aboutBtn.Size = new System.Drawing.Size(82, 34);
            this.aboutBtn.TabIndex = 13;
            this.aboutBtn.Text = "About...";
            this.aboutBtn.UseVisualStyleBackColor = true;
            this.aboutBtn.Click += new System.EventHandler(this.aboutBtn_Click);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer1.Location = new System.Drawing.Point(-1, 51);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.dataGridViewer);
            this.splitContainer1.Panel1.Controls.Add(this.InputText);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.FileSelectionListBox);
            this.splitContainer1.Size = new System.Drawing.Size(943, 424);
            this.splitContainer1.SplitterDistance = 314;
            this.splitContainer1.TabIndex = 14;
            // 
            // dataGridViewer
            // 
            this.dataGridViewer.AllowUserToAddRows = false;
            this.dataGridViewer.AllowUserToDeleteRows = false;
            this.dataGridViewer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewer.Location = new System.Drawing.Point(3, 0);
            this.dataGridViewer.Name = "dataGridViewer";
            this.dataGridViewer.Size = new System.Drawing.Size(311, 424);
            this.dataGridViewer.TabIndex = 7;
            this.dataGridViewer.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewer_CellContentClick);
            // 
            // testbuttn
            // 
            this.testbuttn.Location = new System.Drawing.Point(16, 19);
            this.testbuttn.Name = "testbuttn";
            this.testbuttn.Size = new System.Drawing.Size(194, 23);
            this.testbuttn.TabIndex = 15;
            this.testbuttn.Text = "Load DataGrid";
            this.testbuttn.UseVisualStyleBackColor = true;
            this.testbuttn.Click += new System.EventHandler(this.testbuttn_Click);
            // 
            // testbuttn2
            // 
            this.testbuttn2.Location = new System.Drawing.Point(615, 19);
            this.testbuttn2.Name = "testbuttn2";
            this.testbuttn2.Size = new System.Drawing.Size(91, 23);
            this.testbuttn2.TabIndex = 16;
            this.testbuttn2.Text = "test response";
            this.testbuttn2.UseVisualStyleBackColor = true;
            this.testbuttn2.Click += new System.EventHandler(this.testbuttn2_Click);
            // 
            // tagsBindingSource
            // 
            this.tagsBindingSource.DataMember = "Tags";
            this.tagsBindingSource.DataSource = this.filesListBindingSource;
            // 
            // filesListBindingSource
            // 
            this.filesListBindingSource.DataSource = typeof(Emerson_Excel_Tool.FileStats);
            // 
            // ToolForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(946, 524);
            this.Controls.Add(this.testbuttn2);
            this.Controls.Add(this.testbuttn);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.aboutBtn);
            this.Controls.Add(this.RemoveFilesSelected);
            this.Controls.Add(this.openFilesButton);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.helloWorldLabel);
            this.Controls.Add(this.runExcelProcess);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MinimumSize = new System.Drawing.Size(962, 563);
            this.Name = "ToolForm";
            this.Text = "The Emerson <Vallen Tests> Importer Exporter";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ToolForm_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewer)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tagsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.filesListBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button runExcelProcess;
        private System.Windows.Forms.Label helloWorldLabel;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox InputText;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button openFilesButton;
        private System.Windows.Forms.Button RemoveFilesSelected;
        private System.Windows.Forms.Button aboutBtn;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Button testbuttn;
        private System.Windows.Forms.Button testbuttn2;
        private System.Windows.Forms.DataGridView dataGridViewer;
        private System.Windows.Forms.BindingSource filesListBindingSource;
        private System.Windows.Forms.BindingSource tagsBindingSource;
        private System.Windows.Forms.ListBox FileSelectionListBox;
    }
}