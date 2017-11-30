using System;
using System.Collections;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using System.Collections.Specialized;

namespace Education
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ProcessData();
            Application.Run(new Form1());
        }

        public static int ExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }

        public static void ProcessData()
        {
            string[] files = Directory.GetFiles(@"C:\Users\jrent\Documents\test", "*.xls");
            List<Category> cat = new List<Category>();
            List<Category> mainAndSubCategories = new List<Category>();
            foreach (string file in files)
            {
                Excel.Application excelApp = new Excel.ApplicationClass();
                excelApp.Workbooks.Open(file, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value);

                Excel.Worksheet activeWorksheet = ((Excel.Worksheet)excelApp.Application.ActiveWorkbook.Sheets[1]);

                List<String> annualDates = new List<String>();
                List<String> fiscalDates = new List<String>();
                List<String> schoolYearDates = new List<String>();
                List<String> series = new List<String>();
                List<String> category = new List<String>();

                // Stores the other category and the span of the merged cells




                Excel.Range allCells = (Excel.Range)activeWorksheet.UsedRange;

                foreach (Excel.Range cell in allCells)
                {
                    String cellValue = Convert.ToString(cell.Value);
                    int cellColor = Convert.ToInt32(cell.Font.Color);
                    if (cellValue != null && !cellValue.Equals(""))
                    {
                        if (cellColor == 16776960)
                        {
                            mainAndSubCategories = FindSubCategories(cat, activeWorksheet, cell, cellValue);
                        }
                        if (cellColor == 255)
                        {
                            annualDates.Add(cellValue);
                        }
                        if (cellColor == 16711935)
                        {
                            category.Add(cellValue);
                        }
                        if (cellColor != 0)
                        {
                            series.Add(cellValue);
                        }

                    }

                }
                /*
                List<String> fullCategory = new List<String>();
                if (category.Count > annualDates.Count)
                {
                    foreach (Tuple<String, int> otherCat in otherCategory)
                    {
                        foreach (String cat in category)
                        {
                            fullCategory.Add(cat + otherCat.Key);

                        }
                    }
                }*/
                var i = 0;
                while (i <= mainAndSubCategories.Count)
                {
                    foreach (KeyValuePair<string, int> kvp in mainAndSubCategories[i].MainCategory)
                    {
                        Console.WriteLine("subcategories are" + m);
                    }

                }
            

                excelApp.Quit();
            }
            Console.WriteLine("Done comments");
        }

        private static List<Category> FindSubCategories(List<Category> cat,
                                              Excel.Worksheet activeWorksheet,
                                              Excel.Range cell, string cellValue)
        {
                int mergedNumberOfCells = Convert.ToInt32(cell.Cells.MergeArea.Count);
                int cellRow = Convert.ToInt32(cell.Row);
                int cellColumn = Convert.ToInt32(cell.Column);
                ListDictionary subCatValueAndLength = new ListDictionary();
                Dictionary<String, int> catValueAndLength = new Dictionary<String, int>();
                int[] location = new int[2];
                location[0] = cellColumn;
                location[1] = cellRow;
                var cellBelow = (activeWorksheet.Cells[location[1] + 1, location[0]] as Excel.Range);
                var cellBelowVal = cellBelow.Value;
                var cellAbove = (activeWorksheet.Cells[location[1] - 1, location[0]] as Excel.Range);
                var cellAboveVal = cellAbove.Value;
                var selectedCell = (activeWorksheet.Cells[location[0], location[1]] as Excel.Range);
                // Check if cell below is shaded same color
                if (Convert.ToInt32(cellBelow.Font.Color) == 16776960 && cellBelow.MergeCells.Equals(true))
                {
                    // Find the cells below merged range and if it is divisable into the selected cell
                    String range = cellBelow.MergeArea.Address;
                    var rangeOfSelectedCells = cell.MergeArea.Address;
                    var numberOfColumnsOfCellsBelow = cellBelow.MergeArea.Columns.Count;
                    int numberOfColumnsOfSelectedCells = 0;
                    var colon = rangeOfSelectedCells.IndexOf(':');
                    if (colon != -1)
                    {
                        string startColumn = Convert.ToString(rangeOfSelectedCells[colon - 3]);
                        string endColumn = Convert.ToString(rangeOfSelectedCells[colon + 2]);
                        var sColumn = ExcelColumnNameToNumber(startColumn);
                        var eColumn = 1 + ExcelColumnNameToNumber(endColumn);
                        numberOfColumnsOfSelectedCells = eColumn - sColumn;

                    }

                    if (numberOfColumnsOfSelectedCells % numberOfColumnsOfCellsBelow == 0
                        && numberOfColumnsOfSelectedCells / numberOfColumnsOfCellsBelow != 1)
                    {
                        // determine the number of sub categories under the selected cells
                        int numberOfSubCategories = (numberOfColumnsOfSelectedCells / numberOfColumnsOfCellsBelow);
                        int totalNumberOfColumns = numberOfColumnsOfSelectedCells;
                        // get all the subcategories based on how many and location
                        var i = 0;
                        while (i < totalNumberOfColumns)
                        {
                            int[] rangeOfCells = new int[2];
                            rangeOfCells[0] = location[0] + i;
                            rangeOfCells[1] = location[1] + 1;
                            var subcategorySelection = (activeWorksheet.Cells[rangeOfCells[1], rangeOfCells[0]] as Excel.Range);
                            subCatValueAndLength.Add((String)subcategorySelection.Value, numberOfColumnsOfCellsBelow);
                            i = i + (numberOfColumnsOfCellsBelow);
                        }
                        catValueAndLength.Add(cellValue, numberOfColumnsOfSelectedCells);
                        cat.Add(new Category(catValueAndLength, subCatValueAndLength));
                        var aboveValue = cellAbove.Value;
                        var value = selectedCell.Value;


                    }
                    else if ((numberOfColumnsOfSelectedCells % numberOfColumnsOfCellsBelow == 0
                        && numberOfColumnsOfSelectedCells / numberOfColumnsOfCellsBelow == 1))
                {
                    catValueAndLength.Add(cellValue, numberOfColumnsOfSelectedCells);
                    cat.Add(new Category(catValueAndLength, null));
                }
                    return cat;
                }
                return cat;
        }
    } 
    
}
