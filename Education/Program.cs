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
        public static void ProcessData()
        {
            string[] files = Directory.GetFiles(@"C:\Users\jrent\Documents\test", "*.xls");
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
                List<Tuple<String,int>> otherCategory = new List<Tuple<String,int>>();
                List<Tuple<String, int[]>> cellLocations = new List<Tuple<String, int[]>>();





                Excel.Range allCells = (Excel.Range)activeWorksheet.UsedRange;

                foreach (Excel.Range cell in allCells)
                {
                    String cellValue = Convert.ToString(cell.Value);
                    int mergedNumberOfCells = Convert.ToInt32(cell.Cells.MergeArea.Count);
                    int cellRow = Convert.ToInt32(cell.Row);
                    int cellColumn = Convert.ToInt32(cell.Column);
                    int cellColor = Convert.ToInt32(cell.Font.Color);

                    if (cellValue != null && !cellValue.Equals(""))
                    {
                        if (cellColor == 16776960)
                        {
                            int[] location = new int[2];
                            location[0] = cellRow;
                            location[1] = cellColumn;
                            cell.UnMerge();
                            otherCategory.Add(Tuple.Create(cellValue, mergedNumberOfCells));
                            cellLocations.Add(Tuple.Create(cellValue, location));

                            var mergedArea = cell.Cells.MergeArea;
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
                
                System.Console.WriteLine("Series are: " + series);
                System.Console.WriteLine("Dates are: " + annualDates);
                System.Console.WriteLine("Other Categories are: " + otherCategory);

                excelApp.Quit();
            }
            Console.WriteLine("Done comments");
        }
    }
    
}
