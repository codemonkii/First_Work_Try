using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace sales_tax
{
    public class SalesTax
    {
        /// <summary>
        /// This is a class to perform Sales tax state counts and create new xlsx files for Accounting
        /// </summary>
        static object useDefault = Type.Missing;

        private string tableToTax;
        private string jobnumber;
        private string package = "ALL";
        private string drop = "ALL";
        private string xlsName = string.Empty;
        private string xlsPublicFolder = Properties.Settings.Default.strSalesTaxFolder;
        private string xlsFullName = string.Empty;
        private string stateField;
        private List<string> lstStates;
        private System.Data.DataTable dataTable;

        // Used by Excel
        private Application excel;
        private Workbook excelWorkbook;
        private Worksheet excelWorksheet;

        /// <summary>
        /// Constructor to receive DataTable object with State and Count Columns
        /// DataTable Should be a Distribution of all states
        /// NOTE: Column 0 = state value as string; Column 1 = count value as integer
        /// </summary>
        /// <param name="jobnumber"></param>
        /// <param name="package"></param>
        /// <param name="drop"></param>
        /// <param name="dataTable"></param>
        public void RunSalesTax(string jobnumber, string package, string drop, System.Data.DataTable dataTable)
        {
            try
            {
                // Set up internal Variables
                this.dataTable = dataTable;
                this.jobnumber = jobnumber;
                this.package = package;
                this.drop = drop;
                this.xlsName = jobnumber + ".xlsx";

                this.PopulateStateList();

                this.xlsFullName = Path.Combine(this.xlsPublicFolder, this.xlsName);

                this.removeNulls();

                this.GenerateXls();
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Constructor to receive DataTable object with State and Count Columns
        /// DataTable Should be a Distribution of all states
        /// NOTE: Column 0 = state value as string; Column 1 = count value as integer
        /// This Constructor will have no package and drop passed
        /// </summary>
        /// <param name="jobnumber"></param>
        /// <param name="package"></param>
        /// <param name="drop"></param>
        /// <param name="dataTable"></param>
        public void RunSalesTax(string jobnumber, System.Data.DataTable dataTable)
        {
            try
            {
                // Set up internal Variables
                this.dataTable = dataTable;
                this.jobnumber = jobnumber;
                this.xlsName = jobnumber + ".xlsx";

                this.PopulateStateList();

                this.xlsFullName = Path.Combine(this.xlsPublicFolder, this.xlsName);

                this.removeNulls();

                this.GenerateXls();
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Populating the list of states for which we need counts
        /// </summary>
        private void PopulateStateList()
        {
            try
            {
                string states = Properties.Settings.Default.listOfReportStates;
                this.lstStates = new List<string>();

                // Split method requires character array for the delimiter, so set it up
                char[] deli = {'|'};

                // Split the string delimited with "|" into an array
                string[] stateArray = states.Split(deli);

                // Populate the list with the values 
                for (int i = 0; i < stateArray.Length; i++)
                {
                    this.lstStates.Add(stateArray[i]);
                }
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Private method to do the work
        /// Will create an Exel file whene there is none
        /// Will add a tab when there is one
        /// </summary>
        private void GenerateXls()
        {
            if (!File.Exists(this.xlsFullName))
            {
                // If the file does not exist, create it
                try
                {
                    this.CreateXls();
                }
                catch
                {
                    throw;
                }
            }

            try
            {
                /*
                if (this.sqlConnection.State != ConnectionState.Open)
                {
                    this.sqlConnection.Open();
                }
                */
                
                // Make a new tab name
                string tabName = "p-" + this.package + " d-" + this.drop;

                // Open the excel file for the job
                this.excel = new Application();
                this.excelWorkbook = excel.Workbooks.Open(this.xlsFullName);

                // Find the current tab, if there is one
                int numberOfTabs = this.excel.Sheets.Count;
                int currentTab = 0;

                for (int i = 1; i <= numberOfTabs; i++)
                {
                    if (this.excelWorkbook.Worksheets[i].Name == tabName || this.excelWorkbook.Worksheets[i].Name == "empty")
                    {
                        currentTab = i;
                        break; // if we find a current tab, get out of here
                    }
                }

                // Open the tab that I need to work with (either a brand new tab, or the "current tab")
                if (currentTab == 0)
                {
                    this.excelWorksheet = this.excelWorkbook.Worksheets.Add();
                }
                else
                {
                    this.excelWorksheet = this.excelWorkbook.Worksheets[currentTab];
                }

                this.excelWorksheet.Name = tabName;

                int currRow = 1;

                // Header
                SetCellValue(this.excelWorksheet, "a" + currRow.ToString(), this.jobnumber + " Package " + this.package + " Drop " + this.drop);

                currRow += 2;
                
                // Do our state rows first
                foreach (string state in this.lstStates)
                {
                    // State Labels
                    SetCellValue(this.excelWorksheet, "a" + currRow.ToString(), state);
                    SetCellValue(this.excelWorksheet, "b" + currRow.ToString(), this.GetCount(state));
                    currRow += 1;
                }

                // Now, do the "Other" and "Total" Rows
                SetCellValue(this.excelWorksheet, "a" + currRow.ToString(), "Other");
                SetCellValue(this.excelWorksheet, "b" + currRow.ToString(), this.GetCount("other"));
                currRow += 1;
                SetCellValue(this.excelWorksheet, "a" + currRow.ToString(), "Total");
                SetCellValue(this.excelWorksheet, "b" + currRow.ToString(), this.GetCount("total"));

                this.excelWorksheet.Columns.AutoFit();

            }
            catch
            {
                throw;
            }
            finally
            {
                try
                {
                    // Try to save the workbook, so that the user does not get asked to Save
                    this.excelWorkbook.Save();
                }
                catch { }

                try
                {
                    // Close the workbook
                    this.excelWorkbook.Close(useDefault);
                }
                catch { }

                this.excelWorksheet = null;
                this.excelWorkbook = null;

                try
                {
                    // Quit the Excel Application
                    this.excel.Quit();
                }
                catch { }

                this.excel = null;
            }
        }

        /// <summary>
        /// Method that will create an Excel File
        /// </summary>
        private void CreateXls()
        {
            // Try to create the Excel Document
            try
            {
                this.excel = new Application();
                
                this.excel.SheetsInNewWorkbook = 1;
                this.excelWorkbook = this.excel.Workbooks.Add();
                this.excelWorksheet = excelWorkbook.Worksheets[1];
                this.excelWorksheet.Name = "empty";

                this.excelWorkbook.SaveAs(this.xlsFullName,
                            useDefault, useDefault, useDefault, useDefault, useDefault,
                            XlSaveAsAccessMode.xlNoChange, useDefault, useDefault, useDefault, useDefault, useDefault);
                
            }
            catch
            {
                throw;
            }
            finally
            {                
                try
                {
                    excelWorkbook.Close(useDefault);
                }
                catch 
                { 
                    // nothing 
                }
                try
                {
                    excel.Quit();
                }
                catch
                {
                    // nothing
                }

                this.excelWorksheet = null;
                this.excelWorkbook = null;
                this.excel = null;
            }
        }

        /// <summary>
        /// Method to remove null values from the first column of the data table and set them equal to empty string
        /// </summary>
        private void removeNulls()
        {
            try
            {
                foreach (DataRow row in this.dataTable.Rows)
                {
                    if (row.IsNull(0))
                    {
                        row[0] = string.Empty;
                    }
                }
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Helper to set the cell value
        /// </summary>
        /// <param name="targetSheet"></param>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        private static void SetCellValue(Worksheet targetSheet, string cell, object value)
        {
            targetSheet.get_Range(cell).set_Value(XlRangeValueDataType.xlRangeValueDefault, value);
        }

        /// <summary>
        /// Counts the state values from the Data Table
        /// </summary>
        /// <param name="whatToCount"></param>
        /// <returns></returns>
        private int GetCount(string stateToCount)
        {
            int retValue = 0;

            try
            {
                int count = 0;
                bool skipState;
                switch (stateToCount)
                {
                    case "other":
                        // Total any values that are not one of the states that we report specifically on
                        foreach (DataRow dr in this.dataTable.Rows)
                        {
                            skipState = false;
                            foreach (string state in this.lstStates)
                            {
                                if (dr[0] == null)
                                {
                                    // had problems with null states
                                    // so, we will bypass, they will be counted, however, in the "Other" Total
                                    break;
                                }
                                if ((string)dr[0] == state)
                                {
                                    skipState = true;
                                    break;
                                }
                            }
                                
                            // if the state is not one of our list states, we should total the count up
                            if (skipState)
                            {
                                break;
                            }
                            else
                            {
                                count += (int)dr[1];
                            }
                        }
                        
                        break;
                        
                    case "total":
                        // Total all values in the data table
                        foreach (DataRow dr in this.dataTable.Rows)
                        {
                            count += (int)dr[1];
                        }
                        break;
                        
                    default:
                        // Use LINQ to Query the DataTable
                        // Columns in the data table should be: 
                        // Index 0 = state
                        // Index 1 = count
                        count = (from DataRow dr in this.dataTable.Rows 
                                    where (string)dr[0] == stateToCount
                                    select (int)dr[1]).FirstOrDefault();
                        break;
                }

                retValue = count;
                return retValue;
            }
            catch
            {
                throw;
            }
        }
    }
}
