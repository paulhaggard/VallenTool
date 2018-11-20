﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToolkit
{
    public class ExcelPort
    {
        #region Properties

        /// <summary>
        /// The current excel application
        /// </summary>
        private Excel.Application app { get; set; } = null; //= new Excel.Application();

        /// <summary>
        /// The current workbook open in the excel application
        /// </summary>
        private Excel._Workbook wb { get; set; } = null;

        /// <summary>
        /// The current active sheet in the current workbook
        /// </summary>
        private Excel._Worksheet ws { get; set; } = null;

        /// <summary>
        /// Flag indicating when Excel is open and running
        /// </summary>
        public bool isAppOpen { get; private set; } = false;

        /// <summary>
        /// Flag indicating when a workbook is open in Excel
        /// </summary>
        public bool isWbOpen { get; private set; } = false;

        #endregion

        #region Constructor

        public ExcelPort(bool open = true)
        {
            if (open)
                OpenApp();
        }

        #endregion

        #region Events

        /// <summary>
        /// Ensures that the active sheet will always be the active sheet in this class, triggered any time focus changes from one sheet to another
        /// </summary>
        /// <param name="Sh"></param>
        /// <param name="Target"></param>
        private void App_SheetActivate(object Sh)
        {
            ws = wb.ActiveSheet;
        }

        /// <summary>
        /// Triggers when the current workbook is closed
        /// </summary>
        /// <param name="Wb"></param>
        /// <param name="Cancel"></param>
        private void App_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            isWbOpen = false;
            wb = null;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Attempts to open a new instance of excel
        /// </summary>
        /// <returns>Returns true if successful or if excel is already running, returns false if it failed</returns>
        public bool OpenApp()
        {
            if (isAppOpen)
                return true;
            else
            {
                try
                {
                    app = new Excel.Application();
                    isAppOpen = true;
                    app.SheetActivate += App_SheetActivate;
                    app.WorkbookBeforeClose += App_WorkbookBeforeClose;
                    app.DisplayAlerts = false;  // Suppresses the save prompt after closing excel.
                    return true;
                }
                catch(Exception)
                {
                    // Makes sure to reset the variables in case they might have gotten set before the exception occured
                    app = null;
                    isAppOpen = false;
                    return false;
                }
            }
        }

        /// <summary>
        /// Attempts to close excel
        /// </summary>
        /// <returns>Returns true if successful returns false otherwise</returns>
        public bool CloseApp()
        {
            if (!isAppOpen)
                return true;
            else if (!app.Quitting)
            {
                // Closes the workbook if it's open
                if (isWbOpen)
                    wb.Close(false);

                app.Quit();
                isAppOpen = false;
                return true;
            }
            else
                return false;
        }

        /// <summary>
        /// Opens a workbook in the current excel instance
        /// </summary>
        /// <param name="workbook">name of the workbook to open</param>
        /// <param name="createIfDNE">if true, the function will create the workbook if it does not already exist</param>
        /// <returns>Returns true if successful, returns false otherwise</returns>
        public bool OpenWorkbook(string workbook, bool createIfDNE = true)
        {
            if (isWbOpen)
                throw new InvalidOperationException("You must close all other workbooks before opening a new one");
            if (isAppOpen)
            {
                if (DoesWBExist(workbook))
                {
                    wb = app.Workbooks.Open(workbook);
                    wb.Saved = true;    // Automatically saves the workbook when excel quits
                    isWbOpen = true;
                    return true;
                }

                // Creates the workbook if told to do so
                if (createIfDNE)
                    return CreateWorkbook(workbook);
                return false;
            }
            return false;
        }

        /// <summary>
        /// Creates a new excel workbook with the given name
        /// </summary>
        /// <param name="workbook">name of the workbook to create</param>
        /// <returns>Returns true if a workbook with the same name didn't already exist and it was successful creating one
        /// Returns false otherwise</returns>
        public bool CreateWorkbook(string workbook)
        {
            if (isWbOpen)
                throw new InvalidOperationException("You must close all other workbooks before opening a new one");

            if (DoesWBExist(workbook) || !isAppOpen)
                return false;
            else
            {
                wb = app.Workbooks.Add();
                wb.SaveAs(workbook);
                ws = wb.ActiveSheet;
                isWbOpen = true;
                return true;
            }
        }

        /// <summary>
        /// Checks to see if a workbooks exists in the current excel instance
        /// </summary>
        /// <param name="workbook">name of the workbook to check</param>
        /// <returns>Returns true if the workbook exists, returns false otherwise</returns>
        public bool DoesWBExist(string workbook)
        {
            foreach (Excel.Workbook wb in app.Workbooks)
                if (wb.Name.Equals(workbook))
                    return true;
            return false;
        }

        /// <summary>
        /// Toggles whether the user can see the excel interface or not
        /// </summary>
        /// <param name="value"></param>
        public void setVisible(bool value)
        {
            app.Visible = value;
        }

        #endregion
    }
}
