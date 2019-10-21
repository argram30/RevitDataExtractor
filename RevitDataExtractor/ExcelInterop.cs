using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace RevitDataExtractor
{
    public class Excel
    {
        //ATTRIBUTES
        public Microsoft.Office.Interop.Excel.Application excelApp;
        //the excel application object
        private Workbook activeWorkbook; //the active workbook
        private Sheets worksheets; //the collection of all worksheets in the active workbook
        private Worksheet activeWorksheet; //the active worksheet
        private Range activeRange;
        private string firstCellOfActiveRange;
        private string lastCellofActiveRange;
        public bool selectionMade = false;
        public int ActiveRows;
        public int ActiveColumns;

        //PROPERTIES

        /// <summary>
        /// Returns a list of the names of each worksheet in the active excel document.
        /// </summary>
        public List<string> WorksheetNames
        {
            get
            {
                //a list to hold the strings
                List<string> myNames = new List<string>();

                //loop over the worksheets and populate the list
                foreach (Worksheet s in worksheets)
                {
                    myNames.Add(s.Name);
                }
                return myNames;
            }
        }

        /// <summary>
        /// Returns the full file path of the currently open document.
        /// </summary>
        public string ActiveWorkookName
        {
            get
            {
                return activeWorkbook.FullName;
            }
        }

        /// <summary>
        /// Returns the active worksheet object
        /// </summary>
        public Worksheet ActiveWorksheet
        {
            get
            {
                return activeWorksheet;
            }
            set
            {
                activeWorksheet = value;
            }
        }

        public string ActiveWorksheetName
        {
            get
            {
                return activeWorksheet.Name;
            }
        }

        public string[,] CurrentSelectionData { get; set; }
        public Dictionary<string, string> ActiveSelectionDictionary = new Dictionary<string, string>();

        public Range ActiveRange
        {
            get
            {
                return activeRange;
            }
            set
            {
                activeRange = value;
            }
        }

        public string FirstCellOfActiveRange
        {
            get { return firstCellOfActiveRange; }
            set { firstCellOfActiveRange = value; }
        }

        public string LastCellOfActiveRange
        {
            get { return lastCellofActiveRange; }
            set { lastCellofActiveRange = value; }
        }



        public void open(string filePath)
        {
            activeWorkbook.Close();
            activeWorkbook = excelApp.Workbooks.Open(filePath);
            worksheets = activeWorkbook.Worksheets;
        }

        public void initNew()
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            activeWorkbook = excelApp.Workbooks.Add();
            worksheets = activeWorkbook.Worksheets;

            //disable user prompts
            excelApp.DisplayAlerts = false;


        }

        public void writeRow(string worksheet, int row, List<string> data)
        {
           //set the active worksheet
           activeWorksheet = (Worksheet)worksheets[worksheet];
           ((_Worksheet)activeWorksheet).Activate();
           //create an array to pass into the range
           Object[,] myData = new Object[1, data.Count];
           //loop over the incoming list to populate myData
           for (int i = 0; i < data.Count; i++)
           {
               myData[0, i] = data[i];
           }
           //create a range and set it equal to myData
           Range myStartCell = (Range)activeWorksheet.Cells[row, 1];
           Range myEndCell = (Range)activeWorksheet.Cells[row, data.Count];
           Range myTargetRange = (Range)activeWorksheet.get_Range(myStartCell, myEndCell);
           myTargetRange.Value2 = myData;
       }

       public void writeColumn(int worksheet, int column, List<string> data)
       {
           //set the active worksheet
           activeWorksheet = (Worksheet)worksheets[worksheet];
           ((_Worksheet)activeWorksheet).Activate();
           //create an array to pass into the range
           Object[,] myData = new Object[data.Count, 1];
           //loop over the incoming list to populate myData
           for (int i = 0; i < data.Count; i++)
           {
               myData[i, 0] = data[i];
           }
           //create a range and set it equal to myData
           Range myStartCell = (Range)activeWorksheet.Cells[2, column];
           Range myEndCell = (Range)activeWorksheet.Cells[data.Count + 1, column];
           Range myTargetRange = (Range)activeWorksheet.get_Range(myStartCell, myEndCell);
           myTargetRange.Value2 = myData;
       }
    }
}

