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

        public static int NumberOfColumnsOfMergedCells(String cellAddress)
        {
            var cellAddressString = cellAddress.ToString();
            var colon = cellAddressString.IndexOf(':');
            if (colon != -1)
            {
                string startColumn = Convert.ToString(cellAddressString[colon - 3]);
                string endColumn = Convert.ToString(cellAddressString[colon + 2]);


                if (string.IsNullOrEmpty(startColumn)) throw new ArgumentNullException("columnName");

                startColumn = startColumn.ToUpperInvariant();

                int sum = 0;

                for (int i = 0; i < startColumn.Length; i++)
                {
                    sum *= 26;
                    sum += (startColumn[i] - 'A' + 1);
                }

                if (string.IsNullOrEmpty(endColumn)) throw new ArgumentNullException("columnName");

                endColumn = endColumn.ToUpperInvariant();

                int sum2 = 0;

                for (int i = 0; i < endColumn.Length; i++)
                {
                    sum2 *= 26;
                    sum2 += (endColumn[i] - 'A' + 1);
                }
                int numberOfCols = 0;
                numberOfCols = sum2 - sum;

                return numberOfCols + 1;
            }
            return 0;

        }

        public static String createNeum(String neum)
        {
            var modifiedNeum = Regex.Replace(neum, @"\n+", "");

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
                        if (Char.IsUpper(firstLetter))
                        {
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
            List<Tuple<String, DateTime?, int>> subSubCatsAndDates = new List<Tuple<String, DateTime?, int>>();
            List<DataPoint> dp = new List<DataPoint>();
            List<DateTime> formattedDates = new List<DateTime>();
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


                
                List<String> schoolYearDates = new List<String>();
                List<String> series = new List<String>();
                List<String> category = new List<String>();
                String unformattedNeum = "";

                // Stores the other category and the span of the merged cells


                //ORIENTATION PARAMETER
                //Categories Longer than dates?
                Boolean dateInHeader = true;
                String dateType = "School Year";
                Boolean headerHasMainAndSubCategories = false;
                Boolean headerHasOnlySubCategory = true;

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
                            if(headerHasMainAndSubCategories == false)
                            {
                                mainAndSubCategories = FindSubCategories(cat, activeWorksheet, cell, cellValue, headerHasMainAndSubCategories);
                            }
                            
                        }
                        if (cellColor == 16711680 || cellColor == 255)
                        {
                            if (cellValue != null && cellValue != " ")
                            {
                                formattedDates.Add(FormattedDate(cellValue));
                            }

                        }
                        if (cellColor == 16711935)
                        {
                            category.Add(cellValue);
                            if (headerHasMainAndSubCategories == true)
                            {
                                mainAndSubCategories = FindSubCategories(cat, activeWorksheet, cell, cellValue, headerHasMainAndSubCategories);
                            }
                            
                        }
                        if (cellColor != 0 && cellColor != 16711935 && cellColor != 255
                            && cellColor != 16776960 && cellColor != 16711680)
                        {
                            if(cellValue != null && cellValue != " ")
                            {
                                series.Add(cellValue);
                            }
                            
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
                                    if (dateInHeader == false)
                                    {
                                       subSubCatsAndDates.Add(new Tuple<String, DateTime?, int>(kvp.Key + "_" + subcategory.Key.ToString(), null, counter));
                                    }
                                    else
                                    {
                                        subSubCatsAndDates.Add(new Tuple<String, DateTime?, int>(kvp.Key + "_" + subcategory.Key.ToString(), formattedDates[s], counter));
                                    }
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
                                    if (dateInHeader == true)
                                    {
                                        if (formattedDates.Count != 0)
                                        {
                                            foreach (var d in formattedDates)
                                            {
                                                subSubCatsAndDates.Add(new Tuple<String, DateTime?, int>(kvp.Key, d, counter));
                                            }

                                        }
                                    }
                                    else
                                    {
                                        subSubCatsAndDates.Add(new Tuple<string, DateTime?, int>(kvp.Key, null, counter));
                                    }

                                    length++;
                                    counter++;
                                }

                            }

                        }

                    }
                    i++;
                }
                // add subcategories to main categories
                List<Tuple<String, int>> fullCategoryAndOrder = new List<Tuple<String, int>>();


                //foreach (var mainCat in category)
                if (dateInHeader == true)
                {
                    for (int iter = 0; iter < category.Count; iter++)
                    {
                        String formattedMainCat = category[iter].Replace(".", "");
                        formattedMainCat = formattedMainCat + "_";
                        for (i = 0; i < subSubCatsAndDates.Count; i++)
                        {
                            var oldCategory = subSubCatsAndDates[i].Item1;
                            var someDate = subSubCatsAndDates[i].Item2;
                            dp.Add(new DataPoint(dateType, formattedMainCat + oldCategory, someDate, null, null, subSubCatsAndDates[i].Item3));
                        }
                    }
                }

                //Iterate and add dates to the DataPoint object if they are NOT in the header
                if (dateInHeader == false)
                {
                    int iterator = 0;
                    if (headerHasMainAndSubCategories == false)
                    {
                        foreach (var mainCat in category)
                        {
                            if (subSubCatsAndDates.Count > 0)
                            {
                                foreach (var subCat in subSubCatsAndDates)

                                {
                                    String formattedMainCat = mainCat.Replace(".", "");
                                    formattedMainCat = formattedMainCat + "_";
                                    fullCategoryAndOrder.Add(new Tuple<String, int>(formattedMainCat + subCat.Item1, subCat.Item3));

                                }
                            }
                            else
                            {
                                fullCategoryAndOrder.Add(new Tuple<String, int>(mainCat, iterator));
                                iterator++;
                            }

                        }
                        var duplicateDates = formattedDates.GroupBy(d => d)
                                         .Where(g => g.Count() > 1)
                                         .Select(y => y.Key)
                                         .ToList();
                        if (duplicateDates.Count > 0)
                        {
                            for (var d = 0; d < duplicateDates.Count; d++)
                            //foreach (var fullCategory in fullCategoryAndOrder)
                            {

                                //for (var d = 0; d < duplicateDates.Count; d++)
                                foreach (var fullCategory in fullCategoryAndOrder)
                                {
                                    dp.Add(new DataPoint(dateType, fullCategory.Item1, duplicateDates[d], null, null, fullCategory.Item2));
                                }
                            }
                        }
                        else
                        {
                            for (var date = 0; date < formattedDates.Count; date++)
                            {
                                foreach (var fullCategory in fullCategoryAndOrder)
                                {
                                    dp.Add(new DataPoint(dateType, fullCategory.Item1, formattedDates[date], null, null, fullCategory.Item2));
                                }
                            }

                        }

                    }
                    else
                    {
                        if (subSubCatsAndDates.Count > 0)
                        {
                            foreach (var subCat in subSubCatsAndDates)
                            {
                                fullCategoryAndOrder.Add(new Tuple<String, int>(subCat.Item1, subCat.Item3));
                            }
                        }
                        var duplicateDates = formattedDates.GroupBy(d => d)
                                         .Where(g => g.Count() > 1)
                                         .Select(y => y.Key)
                                         .ToList();
                        if (duplicateDates.Count > 0)
                        {
                            for (var d = 0; d < duplicateDates.Count; d++)
                            //foreach (var fullCategory in fullCategoryAndOrder)
                            {

                                //for (var d = 0; d < duplicateDates.Count; d++)
                                foreach (var fullCategory in fullCategoryAndOrder)
                                {
                                    dp.Add(new DataPoint(dateType, fullCategory.Item1, duplicateDates[d], null, null, fullCategory.Item2));
                                }
                            }
                        }
                        else
                        {
                            for (var date = 0; date < formattedDates.Count; date++)
                            {
                                foreach (var fullCategory in fullCategoryAndOrder)
                                {
                                    dp.Add(new DataPoint(dateType, fullCategory.Item1, formattedDates[date], null, null, fullCategory.Item2));
                                }
                            }
                        }
                    }
                }
                    

                for(var s = 0; s < series.Count; s++)
                {
                    dp[s].Value = series[s];
                }

                //Add neums
                List<Neums> nms = new List<Neums>();
                foreach (var neum in dp)
                {
                    
                    String previousItem1 = "";
                    if (neum.Category != previousItem1)
                    {
                        nms.Add(new Neums(neumAbv + "_" + neum.Category.TrimStart(), unformattedNeum + neum.Category));
                        previousItem1 = neum.Category;
                    }
                    else
                    {
                        previousItem1 = neum.Category;
                    }
                }

                var consolidatedNeums = from n in nms
                                        group n by n.Neum into g
                                        select new { Neum = g.Key, Value = g.Min(d => d.Value) };
                                         

                // WriteXml
                using (StringWriter str = new StringWriter())
                using (XmlTextWriter xml = new XmlTextWriter(str))
                {
                    xml.WriteStartDocument();
                    var worksheetName = activeWorksheet.Name.Replace(" ", "_");
                    xml.WriteStartElement(worksheetName);
                    xml.WriteWhitespace("\n");
                    xml.WriteStartElement("DataSeries");
                    String previousItem = "";
                    foreach (var n in consolidatedNeums)
                    {
                        xml.WriteStartElement("Neum");
                        xml.WriteAttributeString("Full_Neum", n.Neum);
                        xml.WriteAttributeString("Value", n.Value);
                        xml.WriteEndElement();
                        xml.WriteWhitespace("\n");
                    }

                    xml.WriteEndElement();

                    xml.WriteStartElement("DataPoints");
                    foreach (var item in dp)
                    {

                        xml.WriteStartElement(neumAbv);
                        xml.WriteAttributeString("Neum", neumAbv + "_" + item.Category.TrimStart());
                        xml.WriteAttributeString("Category", item.Category);
                        xml.WriteAttributeString("Date", item.Date.ToString());
                        xml.WriteAttributeString("Value", item.Value);
                        xml.WriteAttributeString("PeriodType", item.PeriodType);
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
            int year = 0;
            if (unformattedDate.Contains("-"))
            {
                var splitDate = unformattedDate.Split('-');
                year = Convert.ToInt32(splitDate[0]);
            }
            else
            {
                //Get rid of non digits
                year = Convert.ToInt32(Regex.Replace(unformattedDate, "[^0-9]", ""));
                String yearString = year.ToString();
                if (yearString.Length > 4)
                {
                    year = Convert.ToInt32(yearString.Substring(0,4));
                }

            }

            // Convert to date object
            DateTime dateObject = new DateTime(year, 12, 31);
            // Subtract 1 year and add one day to get last day of December
            return dateObject;

        }


        private static List<Category> FindSubCategories(List<Category> cat,
                                              Excel.Worksheet activeWorksheet,
                                              Excel.Range cell, string cellValue, Boolean headerHasMainAndSubCategories)
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
            // Find the cells below merged range and if it is divisable into the selected cell
            String range = cellBelow.MergeArea.Address;
            var rangeOfSelectedCells = cell.MergeArea.Address;
            var numberOfColumnsOfCellsBelow = cellBelow.MergeArea.Columns.Count;
            int numberOfColumnsOfSelectedCells = 0;
            var colon = rangeOfSelectedCells.IndexOf(':');
            numberOfColumnsOfSelectedCells = NumberOfColumnsOfMergedCells(rangeOfSelectedCells);

            if (headerHasMainAndSubCategories == false)
            {
                // Check if cell below is shaded same color
                if (Convert.ToInt32(cellBelow.Font.Color) == 16776960 && cellBelow.MergeCells.Equals(true))
                {
                    
                    if (numberOfColumnsOfSelectedCells % numberOfColumnsOfCellsBelow == 0
                        && numberOfColumnsOfSelectedCells / numberOfColumnsOfCellsBelow != 1)
                    {
                        // determine the number of sub categories under the selected cells
                        int numberOfSubCategories = (numberOfColumnsOfSelectedCells / numberOfColumnsOfCellsBelow);
                        int totalNumberOfColumns = numberOfColumnsOfSelectedCells;
                        Excel.Range subcategorySelection = null;
                        // get all the subcategories based on how many and location
                        var i = 0;
                        while (i < totalNumberOfColumns)
                        {

                            int[] rangeOfCells = new int[2];
                            rangeOfCells[0] = location[0] + i;
                            // Find subcategory below by adding 1 to the row location
                            rangeOfCells[1] = location[1] + 1;
                            // if location of subcategory has merged cells, go down the rabithole for more subcategories
                            subcategorySelection = (activeWorksheet.Cells[rangeOfCells[1], rangeOfCells[0]] as Excel.Range);
                            String subcategorySelectionValue = "";
                            var subCatAddress = subcategorySelection.MergeArea.Address;
                            var subSubCategoriesNumberOfColumns = NumberOfColumnsOfMergedCells(subCatAddress);
                            if (subSubCategoriesNumberOfColumns == 1)
                            {
                                subcategorySelectionValue = (Convert.ToString((activeWorksheet.Cells[rangeOfCells[1], rangeOfCells[0]] as Excel.Range).Value));
                                subCatValueAndLength.Add(subcategorySelectionValue, numberOfColumnsOfCellsBelow);
                                i = i + (numberOfColumnsOfCellsBelow);
                            }
                            else
                            {
                                var topLevelCategory = (activeWorksheet.Cells[rangeOfCells[1], rangeOfCells[0]] as Excel.Range).Value;
                                var topLevelCatRange = rangeOfCells;
                                topLevelCatRange[1] = topLevelCatRange[1] + 1;
                                for (var s = 0; s < subSubCategoriesNumberOfColumns; s++)
                                {

                                    /// //Adjust range of cells
                                    var subSubCatCol = topLevelCatRange[0] + s;
                                    var subSubCategorySelectionColor = (activeWorksheet.Cells[topLevelCatRange[1], subSubCatCol] as Excel.Range).Font.Color;
                                    var subSubCategorySelectionVal = (activeWorksheet.Cells[topLevelCatRange[1], subSubCatCol] as Excel.Range).Value;

                                    if (Convert.ToInt32(subSubCategorySelectionColor) == 16776960)
                                    {
                                        /// append sub sub categories
                                        subcategorySelectionValue = topLevelCategory + "_" + subSubCategorySelectionVal;
                                    }
                                    else
                                    {
                                        subcategorySelectionValue = topLevelCategory.ToString();
                                    }
                                    subCatValueAndLength.Add(subcategorySelectionValue, numberOfColumnsOfCellsBelow);
                                }
                                //Adjust iterator for the length of the merged cells in the sub sub category
                                i = i + subSubCategoriesNumberOfColumns;
                            }

                        }
                        catValueAndLength.Add(cellValue, numberOfColumnsOfSelectedCells);
                        cat.Add(new Category(catValueAndLength, subCatValueAndLength));
                        var aboveValue = cellAbove.Value;
                        var value = selectedCell.Value;


                    }
                    else if ((numberOfColumnsOfSelectedCells % numberOfColumnsOfCellsBelow == 0
                        && numberOfColumnsOfSelectedCells / numberOfColumnsOfCellsBelow == 1 && cellBelowVal != null))
                    {
                        catValueAndLength.Add(cellValue, numberOfColumnsOfSelectedCells);
                        cat.Add(new Category(catValueAndLength, null));
                    }
                    return cat;
                }
                return cat;       
            }
            if (headerHasMainAndSubCategories == true)
            {
                if (Convert.ToInt32(cellBelow.Font.Color) == 16776960)
                {

                    if (numberOfColumnsOfSelectedCells % numberOfColumnsOfCellsBelow == 0
                        && numberOfColumnsOfSelectedCells / numberOfColumnsOfCellsBelow != 1)
                    {
                        // determine the number of sub categories under the selected cells
                        int numberOfSubCategories = (numberOfColumnsOfSelectedCells / numberOfColumnsOfCellsBelow);
                        int totalNumberOfColumns = numberOfColumnsOfSelectedCells;
                        Excel.Range subcategorySelection = null;
                        // get all the subcategories based on how many and location
                        var i = 0;
                        while (i < totalNumberOfColumns)
                        {

                            int[] rangeOfCells = new int[2];
                            rangeOfCells[0] = location[0] + i;
                            // Find subcategory below by adding 1 to the row location
                            rangeOfCells[1] = location[1] + 1;
                            // if location of subcategory has merged cells, go down the rabithole for more subcategories
                            subcategorySelection = (activeWorksheet.Cells[rangeOfCells[1], rangeOfCells[0]] as Excel.Range);
                            String subcategorySelectionValue = "";
                            var subCatAddress = subcategorySelection.MergeArea.Address;
                            var subSubCategoriesNumberOfColumns = NumberOfColumnsOfMergedCells(subCatAddress);
                            if (subSubCategoriesNumberOfColumns == 0 && subcategorySelection.Columns.Count == 1)
                            {
                                subSubCategoriesNumberOfColumns = 1;
                            }
                            if (subSubCategoriesNumberOfColumns == 1)
                            {
                                subcategorySelectionValue = (Convert.ToString((activeWorksheet.Cells[rangeOfCells[1], rangeOfCells[0]] as Excel.Range).Value));
                                subCatValueAndLength.Add(subcategorySelectionValue, numberOfColumnsOfCellsBelow);
                                i = i + (numberOfColumnsOfCellsBelow);
                            }
                            else
                            {
                                var topLevelCategory = (activeWorksheet.Cells[rangeOfCells[1], rangeOfCells[0]] as Excel.Range).Value;
                                var topLevelCatRange = rangeOfCells;
                                topLevelCatRange[1] = topLevelCatRange[1] + 1;
                                for (var s = 0; s < subSubCategoriesNumberOfColumns; s++)
                                {

                                    /// //Adjust range of cells
                                    var subSubCatCol = topLevelCatRange[0] + s;
                                    var subSubCategorySelection = (activeWorksheet.Cells[topLevelCatRange[1], subSubCatCol] as Excel.Range).Value;
                                    /// append sub sub categories
                                    subcategorySelectionValue = topLevelCategory + "_" + subSubCategorySelection;
                                    subCatValueAndLength.Add(subcategorySelectionValue, numberOfColumnsOfCellsBelow);
                                }
                                //Adjust iterator for the length of the merged cells in the sub sub category
                                i = i + subSubCategoriesNumberOfColumns;
                            }

                        }
                        catValueAndLength.Add(cellValue, numberOfColumnsOfSelectedCells);
                        cat.Add(new Category(catValueAndLength, subCatValueAndLength));
                        var aboveValue = cellAbove.Value;
                        var value = selectedCell.Value;


                    }
                    else if ((numberOfColumnsOfSelectedCells % numberOfColumnsOfCellsBelow == 0
                        && numberOfColumnsOfSelectedCells / numberOfColumnsOfCellsBelow == 1 && cellBelowVal != null))
                    {
                        catValueAndLength.Add(cellValue, numberOfColumnsOfSelectedCells);
                        cat.Add(new Category(catValueAndLength, null));
                    }
                    return cat;
                }
                else
                {
                    catValueAndLength.Add(cellValue, numberOfColumnsOfSelectedCells);
                    cat.Add(new Category(catValueAndLength, null));
                }
            }         
            return cat;
        }
    }

}
