using System;
using System.Collections;
using System.Linq;
using System.Xml.Linq;
using System.Xml;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text.RegularExpressions;
using System.Data;

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

        public static String createNeum(String neum)
        {
            var modifiedNeum = Regex.Replace(neum, @"\s+", " ");
           
            if (String.IsNullOrEmpty(modifiedNeum)) throw new ArgumentNullException("neum");

            var i = 0;
            var first = 0;
            var newNuem = "";
            while (i < modifiedNeum.Length)
            {
                foreach (var letter in modifiedNeum)
                {

                    if (letter == ' ' || letter == modifiedNeum[modifiedNeum.Length - 1])
                    {
                        char firstLetter = modifiedNeum[first];
                        if (Char.IsUpper(firstLetter)){
                            newNuem = newNuem + firstLetter;
                        }
                        if (Char.IsLower(firstLetter))
                        {
                            newNuem = newNuem + Char.ToUpper(firstLetter);
                        }
                        first = i + 1;
                    }
                    i++;
                }
            }
            return Convert.ToString(newNuem);
        }


        public static void ProcessData()
        {
            string[] files = Directory.GetFiles(@"C:\Users\jrent\Documents\test", "*.xls");
            List<Category> cat = new List<Category>();
            List<Category> mainAndSubCategories = new List<Category>();
            List<Tuple<String, DateTime>> subSubCatsAndDates = new List<Tuple<String,DateTime>>();
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
                String unformattedNeum = "";

                // Stores the other category and the span of the merged cells




                Excel.Range allCells = (Excel.Range)activeWorksheet.UsedRange;

                foreach (Excel.Range cell in allCells)
                {
                    String cellValue = Convert.ToString(cell.Value);
                    int cellColor = Convert.ToInt32(cell.Font.Color);
                    if (cellValue != null && !cellValue.Equals(""))
                    {
                        if (cell.Row.Equals(1) && cell.Column.Equals(1))
                        {
                            unformattedNeum = cellValue;
                        }
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
                        if (cellColor != 0 && cellColor != 16711935 && cellColor != 255
                            && cellColor != 16776960)
                        {
                            series.Add(cellValue);
                        }

                    }

                }
                String neumAbv = createNeum(unformattedNeum);

               

                // append sub categories to main categories and interate the correct number of times 
                
                var i = 0;
                
                    
                int counter = 0;
                while (i < mainAndSubCategories.Count)
                {

                    foreach (KeyValuePair<String, int> kvp in mainAndSubCategories[i].MainCategory)
                    {
                        // Create list of Subcategories plus sub sub categories.
                        Console.WriteLine("categories are " + kvp.Key);

                        if (mainAndSubCategories[i].SubCategories != null)
                        {


                            foreach (DictionaryEntry subcategory in mainAndSubCategories[i].SubCategories)
                            {
                                var subCatLength = Convert.ToInt32(subcategory.Value);
                                var mainCatLength = Convert.ToInt32(kvp.Value);
                                var s = 0;

                                while (s < subCatLength)
                                {
                                    //Root.
                                    DateTime formattedDate = FormattedDate(annualDates[counter]);
                                    subSubCatsAndDates.Add(new Tuple<String, DateTime>(kvp.Key + "_" + subcategory.Key.ToString(), formattedDate));
                                    s++;
                                    counter++;
                                }

                            }
                        }
                        if (mainAndSubCategories[i].SubCategories == null && mainAndSubCategories[i].MainCategory != null)
                        {


                            foreach (KeyValuePair<String, int> mc in mainAndSubCategories[i].MainCategory)
                            {
                                var length = 0;
                                while (length < mc.Value)
                                {
                                    DateTime formattedDate = FormattedDate(annualDates[counter]);
                                    subSubCatsAndDates.Add(new Tuple<String, DateTime>(kvp.Key,formattedDate));
                                    length++;
                                    counter++;
                                }

                            }

                        }

                    }
                    i++;
                }



                // add subcategories to main categories
                List<Tuple<String,DateTime>> fullCategoryAndDate = new List<Tuple<String, DateTime>>();
                foreach (var mainCat in category)
                {
                    foreach (var subCat in subSubCatsAndDates) 
                    {
                        String formattedMainCat = mainCat.Replace(".", "");
                        formattedMainCat = formattedMainCat + "_";
                        fullCategoryAndDate.Add(new Tuple<String, DateTime>(formattedMainCat + subCat.Item1, subCat.Item2));
                    }
                }
                List<Tuple<String, DateTime, String>> fullDataPoint = new List<Tuple<String, DateTime, String>>();

                int c = 0;
                foreach (var data in series)
                {
                    fullDataPoint.Add(new Tuple<String, DateTime, String>(fullCategoryAndDate[c].Item1,
                        fullCategoryAndDate[c].Item2, data.ToString()));
                    Console.WriteLine("Data is" + data);
                    c++;
                }
                
                

                // WriteXml
                using (StringWriter str = new StringWriter())
                using (XmlTextWriter xml = new XmlTextWriter(str))
                {
                    xml.WriteStartDocument();
                    var worksheetName = activeWorksheet.Name.Replace(" ", "_");
                    xml.WriteStartElement(worksheetName);
                    xml.WriteWhitespace("\n");
                    xml.WriteStartElement("DataSeries");
                    foreach(var n in fullDataPoint)
                    {
                        xml.WriteStartElement("Neum");
                        xml.WriteAttributeString("Full_Neum", neumAbv + "_" + n.Item1.TrimStart());
                        xml.WriteAttributeString("Value", unformattedNeum + n.Item1);
                        xml.WriteEndElement();
                        xml.WriteWhitespace("\n");
                    }
                    xml.WriteEndElement();
                    
                    xml.WriteStartElement("DataPoints");
                    foreach (var item in fullDataPoint)
                    {
                        
                        xml.WriteStartElement(neumAbv);
                        xml.WriteAttributeString("Neum", neumAbv + "_" + item.Item1.TrimStart());
                        xml.WriteAttributeString("Category", item.Item1);
                        xml.WriteAttributeString("Date", item.Item2.ToString());
                        xml.WriteAttributeString("Value", item.Item3);
                        xml.WriteAttributeString("PeriodType", "School Year");
                        xml.WriteEndElement();
                        xml.WriteWhitespace("\n");
                    }

                    xml.WriteEndElement();
                    xml.WriteEndDocument();

                    

                    // Result is a string.
                    string result = str.ToString();
                    File.WriteAllText(worksheetName + ".xml", result);
                    Console.WriteLine("Length: {0}", result.Length);
                    Console.WriteLine("Result: {0}", result);
                }
                excelApp.Quit();
            }
        }

        private static DateTime FormattedDate(String unformattedDate)
        {
            
            //Get rid of non digits
            String noLetters = Regex.Replace(unformattedDate, "[^0-9]", "");
            //Convert to int
            int year = Convert.ToInt32(noLetters);
            // Convert to date object
            DateTime dateObject = new DateTime(year, 12, 31);
            // Subtract 1 year and add one day to get last day of December
            return dateObject;

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
