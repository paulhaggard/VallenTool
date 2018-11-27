using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Xml;
using System.Xml.Serialization;

namespace Emerson_Excel_Tool
{
    public partial class ToolForm
    {
        #region ListBox Actions: Remove, Multifile add, Empty Listbox

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
        /// Action to multi-select .txt files for processing. 
        /// To be displayed in list box.
        /// </summary>
        private void FileSelectionHelper()
        {
            //DialogResult dr = this.openFileDialog1.ShowDialog();
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                // Read the files
                foreach (String file in openFileDialog1.FileNames)
                {
                    int i = 0;
                    if (!FileSelectionListBox. .Contains(file))
                    {    // Create a List Item.
                        try
                        {
                            var newfile = new FileStats { FileFullPath = file, FileName = "File " + i };
                            FileSelectionListBox.Items.Add(newfile);
                            i++;
                            FileSelectionListBox.TopIndex.ToString();
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
            
        }



        /// <summary>
        /// To prevent the files list from being executed on twice without emptying list,
        /// this action clears the list of file locations after the Excel processing is requested.
        /// </summary>
        private void EmptyTheFileList()
        {
            appFileList.Clear();
        }

        #endregion

        #region Form Configurators
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


        #endregion

    }
}
