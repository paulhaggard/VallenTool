﻿using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Emerson_Excel_Tool
{
    public class DataSet_Processing
    {
        /// <summary>
        /// Data set of the vallen log
        /// </summary>
        public DataSet _vallenLogs { get; private set; } = new DataSet("Vallen Logs");

        //private string _tableName;
        //private string _tableFileLocation;
        public DataSet_Processing()  //initializes an instance of the class
        {

        }



        //Test Variables
        private static string testFileLocation = @"C:\Projects\2007 - Emerson AE\08. Testing\Paul Report Writing\Vallen Sensor Tests\296 a-b\VS900-RIC - oil - vallen reset.r216.txt";
        public List<string> testFileList = new List<string>();


        #region Data Set & Data Tables Code

        //set the working table file name and file path during processing of each file
        public string tableName { get; set; }
        public string tableFileLocation { get; set; }


        ///This section includes the code for implementing the data set and its tables.
        ///

        //Creates a DataTable from input method CreateDataTableFromFile
        //Converts From DataTable object to String Array for a 2D array of any size.
        //Outputs 
        public void GetTableData(out DataTable dt, out string[,] results)
        {
            // v This is unnecessary and can lead to infinite loops, don't use recursive classes they're a mess to work with
            //DataSet_Processing instance = new DataSet_Processing(); /////GOT HUNGUP HERE BADLY! READ MORE.

            // Just use the function from this class instance, or make it a static method
            dt = CreateDataTableFromFile(tableName, tableFileLocation);  //TODO come back and use name field // ERROR

            results = new string[dt.Rows.Count, dt.Columns.Count];

            for (int index = 0; index < dt.Rows.Count; index++)
            {
                if (index == 0) { results[0, 0] = tableName; }
                else
                {
                    for (int columnIndex = 0; columnIndex < dt.Columns.Count; columnIndex++)
                    { results[index, columnIndex] = dt.Rows[index][columnIndex].ToString(); }
                }
            }
        }




        /// <summary>
        /// Create a data table from a file.
        /// </summary>
        /// <returns>Returns the DataTable object that was created from the file</returns>
        public DataTable CreateDataTableFromFile(string _name, string fileLocation)
        {


            DataTable newTable = new DataTable();
            newTable = _vallenLogs.Tables.Add(_name);

            DataColumn FreqColumn =
            newTable.Columns.Add("Frequency", typeof(string));
            newTable.Columns.Add("Response", typeof(string));
            newTable.PrimaryKey = new DataColumn[] { FreqColumn };
            DataRow dr;


            StreamReader sr = new StreamReader(fileLocation);
            string input;
            while ((input = sr.ReadLine()) != null)
            {
                if (input == string.Empty)
                {
                    continue;
                }
                else
                {

                    string[] s = input.Split(new char[] { '\t' });

                    if (s.Length < 2)
                        throw new InvalidDataException("The data file must be formatted with tab-delimitted data containing frequency data and response data!");

                    dr = newTable.NewRow();
                    dr["Frequency"] = s[0];
                    dr["Response"] = s[1];  // ERROR
                    newTable.Rows.Add(dr);
                }
            }
            sr.Close();
            return newTable;
        }


        /// <summary>
        /// DataTable for List
        /// </summary>
        /// <param name="fileLocation"></param>
        /// <returns></returns>
        public DataTable CreateDataTableFromFile(List<string> fileLocation)
        {



            DataTable newtable = new DataTable();
            DataColumn dc;
            DataRow dr;
            for (int i = 0; i < fileLocation.Count; i++)
            {
                dc = new DataColumn();
                dc.DataType = System.Type.GetType("System.String");
                dc.ColumnName = "c1";
                dc.Unique = false;
                newtable.Columns.Add(dc);
                dc = new DataColumn();
                dc.DataType = System.Type.GetType("System.String");
                dc.ColumnName = "c2";
                dc.Unique = false;
                newtable.Columns.Add(dc);

                StreamReader sr = new StreamReader(fileLocation[i]);
                string input;
                while ((input = sr.ReadLine()) != null)
                {
                    if (input == string.Empty)
                    {
                        continue;
                    }
                    else
                    {

                        string[] s = input.Split(new char[] { '\t' });
                        dr = newtable.NewRow();
                        dr["c1"] = s[0];
                        dr["c2"] = s[1];
                        newtable.Rows.Add(dr);
                    }
                }
                sr.Close();
            }
            return newtable;
        }




        #endregion
    }
}




/*
            public DataSet_Processing(string tableName, string fileLocation)
            {
                string _tableName = tableName;
                string _fileLocation = fileLocation;
            
            }
            */
