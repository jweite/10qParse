using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AngleSharp;
using System.Text.RegularExpressions;
using System.Diagnostics;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace AngleSharpTest1
{
    class Program
    {
        static void ExtractFromFiling(string sourceFileName, string ticker, string filingType)
        {
            string extractionsJson = File.ReadAllText(@"config\extractions.json");
            JObject config = JObject.Parse(extractionsJson);

            var companyConfig = config[ticker];
            if (companyConfig == null)
            {
                Console.WriteLine("No config found for ticker " + ticker);
                return;
            }

            var filingConfig = companyConfig["Filings"][filingType];
            if (filingConfig == null)
            {
                Console.WriteLine("No filing config found for filing " + filingType + " under ticker " + ticker);
                return;
            }

            string sourceFilesBaseDirectory = @"C:\temp\Wikipedia_SP500_10Q_Downloader_Output\";
            string sourceFileFullyQualified = sourceFilesBaseDirectory + sourceFileName;
            AngleSharp.Dom.Html.IHtmlDocument htmlDoc = null;
            try
            {
                htmlDoc = ReadAndParseHtmlFile(sourceFileFullyQualified);
            }
            catch(Exception e)
            {
                Console.WriteLine("Exception reading " + sourceFileFullyQualified + ":");
                Console.WriteLine(e);
                return;
            }

            string outputFilesBaseDirectory = @"c:\temp\10QParserOutput\";
            foreach (JObject statement in filingConfig.Value<JArray>())
            {
                foreach (JProperty prop in statement.Properties())
                {
                    string statementTitle =  prop.Value["StatementTitle"].ToString();

                    int statementOccurrenceIndex = (int)prop.Value["StatementOccurrence"];

                    string outputFile = outputFilesBaseDirectory + ticker + "_" + statementTitle + "_" + statementOccurrenceIndex + ".txt";

                    var jLandmarks = ((JArray)prop.Value["Landmarks"]).ToArray<JToken>();
                    List<string> arLandmarks = new List<string>();
                    foreach (var jLandmark in jLandmarks)
                    {
                        arLandmarks.Add(((JValue)jLandmark).ToString());
                    }

                    IDictionary<string, string> rowHeadOverrideDict = new Dictionary<string, string>();
                    IDictionary<string, string> parameterDict = new Dictionary<string, string>();

                    var jOptions = ((JObject)prop.Value["Options"]);
                    if (jOptions != null)
                    {
                        var jRowHeadOverrides = (JArray)(jOptions["RowHeadOverrides"]);
                        if (jRowHeadOverrides != null)
                        {
                            foreach (var jRowHeadOverride in jRowHeadOverrides.Children<JObject>())
                            {
                                var k = jRowHeadOverride.Properties().First<JProperty>().Name;
                                var v = jRowHeadOverride.Properties().First<JProperty>().Value;
                                rowHeadOverrideDict.Add(k, v.ToString());
                            }
                        }

                        var jParameters = (JArray)(jOptions["Parameters"]);
                        if (jParameters != null)
                        {
                            foreach (var jParameter in jParameters.Children<JObject>())
                            {
                                var k = jParameter.Properties().First<JProperty>().Name;
                                var v = jParameter.Properties().First<JProperty>().Value;
                                parameterDict.Add(k, v.ToString());
                            }
                        }
                    }

                    ExtractTableFromHTML(htmlDoc, arLandmarks.ToArray(), statementTitle, statementOccurrenceIndex, outputFile, rowHeadOverrideDict, parameterDict);
                }
            }




        }


        static void Main(string[] args)
        {
            ExtractFromFiling("MMM_mmm-20170630x10q.htm", "MMM","10-Q");

            //ExtractTableFromHTML(@"c:\temp\de-20170430x10q.htm", "STATEMENT OF CONSOLIDATED INCOME", 1, @"c:\temp\deere_consolidated_income_3mon", new Dictionary<string, string> { { "UndifferentiatedTotalAssociatesWithPrecedingBoldHeading", "true" } });
            //ExtractTableFromHTML(@"c:\temp\de-20170430x10q.htm", "STATEMENT OF CONSOLIDATED COMPREHENSIVE INCOME", 1, @"c:\temp\deere_consolidated_comprehensive_income_3mon", new Dictionary<string, string> { { "UndifferentiatedTotalAssociatesWithPrecedingBoldHeading", "true" } });
            //ExtractTableFromHTML(@"c:\temp\de-20170430x10q.htm", "STATEMENT OF CONSOLIDATED INCOME", 2, @"c:\temp\deere_consolidated_income_6mon", new Dictionary<string, string> { { "UndifferentiatedTotalAssociatesWithPrecedingBoldHeading", "true" } });
            //ExtractTableFromHTML(@"c:\temp\de-20170430x10q.htm", "STATEMENT OF CONSOLIDATED COMPREHENSIVE INCOME", 2, @"c:\temp\deere_consolidated_comprehensive_income_6mon", new Dictionary<string, string> { { "UndifferentiatedTotalAssociatesWithPrecedingBoldHeading", "true" } });
            //ExtractTableFromHTML(@"c:\temp\de-20170430x10q.htm", "CONDENSED CONSOLIDATED BALANCE SHEET", 1, @"c:\temp\deere_balance", null);
            //ExtractTableFromHTML(@"c:\temp\de-20170430x10q.htm", "STATEMENT OF CONSOLIDATED CASH FLOWS", 1, @"c:\temp\deere_cashflow_6mon", new Dictionary<string, string> { { "UndifferentiatedTotalAssociatesWithPrecedingBoldHeading", "true" } });

            //ExtractTableFromHTML(@"c:\temp\cat_10qx6302017.htm", "Consolidated Statement of Results of Operations", 1, @"c:\temp\cat", null);
            //ExtractTableFromHTML(@"c:\temp\a17-13367_110q.htm", "CONSOLIDATED STATEMENT OF EARNINGS", 1, @"c:\temp\ibm", null);
            //ExtractTableFromHTML(@"c:\temp\a10-qq32017712017.htm", "CONDENSED CONSOLIDATED STATEMENTS OF OPERATIONS (Unaudited)", 1, @"c:\temp\apple", null);
            //ExtractTableFromHTML(@"c:\temp\mdt-2015q3x10q.htm", "CONDENSED CONSOLIDATED STATEMENTS OF EARNINGS", 1, @"c:\temp\medtronic", null);
            //ExtractTableFromHTML(@"c:\temp\amzn-20170630x10q.htm", "CONSOLIDATED STATEMENTS OF OPERATIONS", 1, @"c:\temp\amazon", null);

            //Console.Write("Done. Press enter to exit");
            //Console.ReadLine();
        }

        public static void ExtractTableFromHTML(AngleSharp.Dom.Html.IHtmlDocument doc, string[] landmarks, string statementTitle, int statementInstanceIndex, string outputPath, IDictionary<string, string> rowHeadOverrides, IDictionary<string, string> config)
        {
            // Select all the DOM's elements
            var tableSelector = doc.QuerySelectorAll("*");

            // Find the sequential set of landmarks that skip past any undesired occurrences of the statement title before the statement itself (ie in the TOC)
            int lastLandmarkIndex = findLandmarks(tableSelector, 0, landmarks);

            // Skip past the last landmark found
            lastLandmarkIndex = skipPastLandmark(tableSelector, lastLandmarkIndex, landmarks[landmarks.Length - 1]);

            // Find the actual statement title landmark
            lastLandmarkIndex = findLandmark(tableSelector, lastLandmarkIndex, statementTitle);                         

            // See if that landmark is contained by a table (assumed, then, to be the statement table)
            int statementTableIndex = findContainingElementByType(tableSelector, lastLandmarkIndex, "table");
            if (statementTableIndex == -1)
            {
                // Landmark is not contained in a table, so statement table is assumed to be first table following the landmark.
                statementTableIndex = findFollowingElementByType(tableSelector, lastLandmarkIndex, "table");
            }
            if (statementTableIndex == -1)
            {
                Console.WriteLine("No landmarked table found");
                return;
            }
            var statementTable = tableSelector[statementTableIndex];


            // Parse statement table just found into a 2d matrix (actually a list of TableRow objects, which contain a list of TableCell objects)
            var tableData = new List<TableRow>();
            var rowElements = statementTable.QuerySelectorAll("TR");
            foreach (var rowElement in rowElements)
            {
                var rowData = new TableRow();

                var colElements = rowElement.QuerySelectorAll("TD");
                foreach (var cellElement in colElements)
                {
                    // Extract cell value
                    TableCell tableCell = new TableCell(cellElement);

                    // Duplicate cell value across all cols it spans.
                    int colSpan = 1;
                    var colAttrs = cellElement.Attributes;
                    var colSpanAttr = colAttrs.Where(a => a.Name == "colspan");     
                    if (colSpanAttr.Count() > 0)
                    {
                        string sColSpan = colSpanAttr.First().Value;
                        Int32.TryParse(sColSpan, out colSpan);
                    }
                    for (int j = 0; j < colSpan; ++j)
                    {
                        rowData.AddCell(tableCell);
                    }
                }

                tableData.Add(rowData);
            }

            // For diagnostic purposes save a csv of the parsed table 
//            writeTableToFile(outputPath + ".tbl", tableData);

            // Extract the column Headings from the table into list
            IList<string> columnHeadings = ExtractColumnHeadings(tableData);
            if (columnHeadings == null)
            {
                Console.WriteLine("FATAL: Cannot find any qualifying heading rows in table");
            }

            // Post-process the rowheads in the table, calculating its relative indentation level (compared to the rest of the rowheads)
            CalcRowheadIndentationLevels(tableData);

            // Post-process the rowheads, linking each to any parents it has, based on rules and clues.
            BuildComplexRowHeads(tableData, rowHeadOverrides, config);

            // Flatten matrix to a list of tuples using some rules
            List <FlattenedRow> results = new List<FlattenedRow>();
            string attributeName = "";

            for (int iRow = 0; iRow < tableData.Count; ++iRow) {
                var row = tableData[iRow];
                int nCols = row.Cells.Count;

                // Create the attribute name for this row from its row head and those of its parents
                attributeName = row.RowHead.Text;
                TableRow row2 = row.parentRow;
                while (row2 != null && row2.RowHead.Text.Length > 0)
                {
                    attributeName = row2.RowHead.Text + "|" + attributeName;
                    row2 = row2.parentRow;
                }
                if (attributeName.Length == 0) continue;    // Assumption: rows without attribute names should be skipped.

                // Standardize attribute name format: only single spaces between words
                attributeName = ConsolidateWhitespace(attributeName);

                // Scan columns
                string colContent = "";
                for (int iCol = 1; iCol < nCols; ++iCol)
                {
                    var col = row.Cells[iCol];
                    colContent = col.Text;

                    // Process numeric columns only.  Exclude centered (heading) cols.
                    if (col.HorizontalAlignment == TableCell.HORIZONTAL_ALIGNMENT.CENTER || !Regex.IsMatch(colContent, @"\(?\d+\)?"))
                    {
                        continue;
                    }

                    // Convert (xxx) to -xxx.  Drop comma separators
                    colContent = colContent.Replace('(', '-').Replace(")", "").Replace(",", "");

                    // Get the heading for this col
                    string heading = columnHeadings[iCol];

                    // Create the tuple with this data and add it to the results list.
                    FlattenedRow flatRow = new FlattenedRow(attributeName, heading, colContent);
                    results.Add(flatRow);
                }
            }

            // HTML column spans can result in duplicated entries: get rid of them.
            var distinctFlatRows = results.Distinct();

            // Write the flattened matrix out to file
            using (StreamWriter fsw = File.CreateText(outputPath))
            {
                foreach (var record in distinctFlatRows)
                {
                    fsw.Write(record + "\r\n");
                }
            }
        }

        public static AngleSharp.Dom.Html.IHtmlDocument ReadAndParseHtmlFile(string path)
        {
            // Set up angleSharp html parser
            var angleSharConfig = Configuration.Default.WithCss();
            var parser = new AngleSharp.Parser.Html.HtmlParser(angleSharConfig);

            // Read the source html file
            FileStream fs = File.OpenRead(path);

            // Parse it into a DOM with AngleSharp
            var doc = parser.Parse(fs);

            return doc;
        }


        public static int findLandmark(AngleSharp.Dom.IHtmlCollection<AngleSharp.Dom.IElement> elementsToScan, int startingIndex, string landmark)
        {
            for (int iElement = startingIndex; iElement < elementsToScan.Length; ++iElement)
            {
                var element = elementsToScan[iElement];
                if (element.TextContent.Trim() == landmark)
                {
                    return iElement;
                }
            }
            return -1;  // Not found
        }

        public static int findLandmarks(AngleSharp.Dom.IHtmlCollection<AngleSharp.Dom.IElement> elementsToScan, int startingIndex, string[] landmarks)
        {
            if (landmarks == null || landmarks.Length == 0) return startingIndex;

            for (int iLandmark = 0; iLandmark < landmarks.Length; ++iLandmark)
            {
                string landmark = landmarks[iLandmark];
                startingIndex = findLandmark(elementsToScan, startingIndex, landmark);
                if (startingIndex == -1)
                {
                    return -1;      // Could not find one of the landmarks
                }
                else if (iLandmark != landmarks.Length - 1) // Don't skip past the last landmark found
                {
                    startingIndex = skipPastLandmark(elementsToScan, startingIndex, landmark);
                }
            }
            return startingIndex;
        }

        public static int findFollowingElementByType(AngleSharp.Dom.IHtmlCollection<AngleSharp.Dom.IElement> elementsToScan, int startingIndex, string elementType)
        {
            for (int iElement = startingIndex; iElement < elementsToScan.Length; ++iElement)
            {
                var element = elementsToScan[iElement];
                if (element.TagName.ToLower() == elementType.ToLower())
                {
                    return iElement;
                }
            }
            return -1;  // Not found
        }

        public static int findContainingElementByType(AngleSharp.Dom.IHtmlCollection<AngleSharp.Dom.IElement> elementsToScan, int startingIndex, string elementType)
        {
            var element = elementsToScan[startingIndex];

            while (element != null && element.TagName.ToLower() != elementType.ToLower())
            {
                element = element.ParentElement;
            }
            if (element != null && element.TagName.ToLower() == elementType.ToLower())
            {
                for (int iElement = 0; iElement < elementsToScan.Length; ++iElement)
                {
                    // Converting back to index numbers is fugly.  
                    if (elementsToScan[iElement] == element)
                    {
                        return iElement;
                    }
                }
                Debug.Assert(false, "findContainingElementByType could not find element known to be in elementsToScan");
                return -1;  // Won't happen.
            }
            return -1;      // No containing element of this type found.
        }

        public static int skipPastLandmark(AngleSharp.Dom.IHtmlCollection<AngleSharp.Dom.IElement> elementsToScan, int startingIndex, string landmark)
        {
            int iElement;
            for (iElement = startingIndex; iElement < elementsToScan.Length && elementsToScan[iElement].TextContent.Trim() == landmark; ++iElement)
                ;
            return iElement;
        }

        public static bool GetConfigBool(string key, IDictionary<string, string>configDict)
        {
            if (configDict == null) return false;
            if (!configDict.ContainsKey(key)) return false;
            if (configDict[key].ToLower() == "true") return true;
            return false;
        }

        public static void CalcRowheadIndentationLevels(List<TableRow> tableData)
        {
            // Find rowHead distinct indentation levels
            var distinctIndentationLevels = tableData
                .GroupBy(row => row.RowHead.Indentation)
                .Select(group => group.First().RowHead.Indentation)
                .OrderBy(indentationLevel => indentationLevel);

            int indentatationLevel = 0;
            Dictionary<double, int> distinctIndentationLevelDict = new Dictionary<double, int>();
            foreach (double d in distinctIndentationLevels)
            {
                distinctIndentationLevelDict.Add(d, indentatationLevel++);
            }

            foreach (var row in tableData)
            {
                row.RowHead.IndentationLevel = distinctIndentationLevelDict[row.RowHead.Indentation];
            }
        }

        public static void BuildComplexRowHeads(List<TableRow> tableData, IDictionary<string, string> rowHeadOverrides, IDictionary<string, string> config)
        {
            for (int iRow = 0; iRow < tableData.Count; ++iRow)
            {
                var row = tableData[iRow];

                // If this rowhead is overridden through config replace it as specified. It stands alone (though the substitution rowhead may|be|complex.
                if (rowHeadOverrides.ContainsKey(row.RowHead.Text))
                {
                    row.RowHead.Text = rowHeadOverrides[row.RowHead.Text];
                    row.parentRow = null;
                    continue;
                }

                // If configured to, associate undifferentiated "Total" headings with the preceding bold heading (ie Deere Income)
                bool foundMatch = false;
                if (row.RowHead.Text == "Total" && GetConfigBool("UndifferentiatedTotalAssociatesWithPrecedingBoldHeading", config))
                {
                    for (int iRow2 = iRow - 1; iRow2 >= 0 && tableData[iRow2].RowHead.Text.Length > 0; --iRow2)
                    {
                        var precedingRow = tableData[iRow2];
                        if (precedingRow.RowHead.Bold)
                        {
                            row.parentRow = precedingRow;
                            foundMatch = true;
                            break;
                        }
                    }                  
                    if (foundMatch)
                    {
                        continue;
                    }
                }

                // If this rowhead is of the form "Total XXXX" and XXX is a preceeding rowhead then that preceding rowhead is its parent (regardless of indentation)
                string s = row.RowHead.Text.ToLower();
                Match m = Regex.Match(s, @"total (.*)");
                if (m.Success)
                {
                    foundMatch = false;
                    string lookFor = m.Groups[1].Captures[0].Value.ToLower();
                    for (int iRow2 = iRow-1; iRow2 >= 0 /* && tableData[iRow2].RowHead.Text.Length > 0 */; --iRow2)
                    {
                        var precedingRow = tableData[iRow2];
                        if (precedingRow.RowHead.Text.ToLower() == lookFor)
                        {
                            row.parentRow = tableData[iRow2];
                            foundMatch = true;
                            break;
                        }
                    }
                    if (foundMatch == true)
                    {
                        continue;
                    }
                }


                // If this rowhead is indented make the first preceeding row at a less indentation level its parent.
                if (row.RowHead.IndentationLevel > 0)
                {
                    for (int iRowInner = iRow - 1; iRowInner >= 0; --iRowInner)
                    {
                        if (row.RowHead.Text == "Total")
                        {
                            // An undifferentiatated "Total" may be the child of a head at its own indentation level.
                            if (tableData[iRowInner].RowHead.IndentationLevel <= row.RowHead.IndentationLevel)
                            {
                                row.parentRow = tableData[iRowInner];
                                break;
                            }
                        }
                        else
                        {
                            if (tableData[iRowInner].RowHead.IndentationLevel < row.RowHead.IndentationLevel)
                            {
                                row.parentRow = tableData[iRowInner];
                                break;
                            }
                        }
                    }
                }
                // If not indented, not bold and has content make previous bold row without content its parent.
                else if (!row.RowHead.Bold /* && row.bRowCellsHaveContent */)
                {
                    for (int iRowInner = iRow - 1; iRowInner >= 0; --iRowInner)
                    {
                        var possibleParent = tableData[iRowInner];
                        if (possibleParent.RowHead.Bold)
                        {
                            if (!possibleParent.bRowCellsHaveContent)
                            {
                                row.parentRow = possibleParent;
                                break;
                            }
                            else
                            {
                                break;      // If the preceeding bold row DOES have content it's not a parent, and curr row is considered to stand alone.
                            }
                        }
                    }
                }
            }
        }



        public static List<string> ExtractColumnHeadings(List<TableRow> tableData)
        {
            // Find first heading row
            int iRow = 0;
            for (iRow = 0; iRow < tableData.Count; ++iRow)
            {
                var row = tableData[iRow];
                bool foundCenteredCell = false;
                foreach (var cell in row.Cells)
                {
                    if (cell.HorizontalAlignment == TableCell.HORIZONTAL_ALIGNMENT.CENTER)
                    {
                        foundCenteredCell = true;
                        break;
                    }
                }
                if (foundCenteredCell)
                {
                    break;
                }
            }
            if (iRow >= tableData.Count)
            {
                return null;
            }

            // For each cell in the first heading row, build the full heading from the centered cells below it and add to list.
            List<string> headings = new List<string>();
            bool centeredCellFound = false;
            for (int iCol = 0; iCol < tableData[iRow].Cells.Count; ++iCol)
            {
                string heading = "";
                for (int ir = iRow; ir < tableData.Count; ++ir)
                {
                    TableCell cell = tableData[ir].Cells[iCol];
                    if (cell.Text.Length > 0)
                    {
                        if (cell.HorizontalAlignment == TableCell.HORIZONTAL_ALIGNMENT.CENTER)
                        {
                            if (heading.Length > 0)
                            {
                                heading += " ";
                            }
                            heading += cell.Text;
                            centeredCellFound = true;
                        }
                        else
                        {
                            if (centeredCellFound)
                            {
                                break;
                            }
                        }
                    }
                }

                // MAKE SURE THIS IS GENERALLY APPLICABLE!!!
                // When headings are just a year then the actual time range they represent is often elsewhere in the table.
                if (heading.StartsWith("20") && heading.Length == 4)
                {
                    // Look for more detail about what these points in time mean elsewhere in the table.  See Deere 10Q.
                    foreach (var row2 in tableData)
                    {
                        string rowHeadTextLower = row2.RowHead.Text.ToLower();

                        // Assumption for now: these explanations are in column 0
                        Match mMonths = Regex.Match(rowHeadTextLower, @"(\w*) months ended");
                        if (mMonths.Success)
                        {
                            string numMonths = mMonths.Groups[1].Captures[0].Value;
                            numMonths = numMonths.Substring(0, 1).ToUpper() + numMonths.Substring(1);

                            Match mDate = Regex.Match(rowHeadTextLower, @"(\w+) (\d+)?, " + heading);
                            if (mDate.Success)
                            {
                                string month = mDate.Groups[1].Captures[0].Value;
                                month = month.Substring(0, 1).ToUpper() + month.Substring(1);
                                string dayOfMonth = mDate.Groups[2].Captures[0].Value;

                                heading = numMonths + " months ended " + month + " " + dayOfMonth + ", " + heading;
                            }
                        }
                    }
                }

                heading = NormalizeTimeRangeHeading(heading);

                headings.Add(heading);
            }
            return headings;
        }

        public static string ConvertToTitleCase(string s)
        {
            char prevChar = ' ';
            StringBuilder builder = new StringBuilder();

            foreach (char c in s.ToCharArray())
            {
                if (prevChar == ' ' || prevChar == '\t' || prevChar == '\n' || prevChar == '\r')
                {
                    builder.Append(c.ToString().ToUpper().ToCharArray()[0]);
                }
                else
                {
                    builder.Append(c);
                }
                prevChar = c;
            }

            return builder.ToString();
        }

        public static string NormalizeTimeRangeHeading(string heading)
        {
            string result = heading;

            Match mMonths = Regex.Match(heading.ToLower(), @"(\w+) months end");    // Allows for ended and ending...
            if (mMonths.Success)
            {
                Match mDate = Regex.Match(heading, @"(\w+)\s+(\d+),?\s+20(\d\d)"); // Month Day, Year
                if (mDate.Success)
                {
                    result = mMonths.Groups[1].Captures[0].Value + " months ending " + mDate.Groups[1].Captures[0].Value + " " + mDate.Groups[2].Captures[0].Value + ", 20" + mDate.Groups[3].Captures[0].Value;
                    return ConvertToTitleCase(result);
                }
            }
            return result;  // If we can't find the expected elements do nothing.
        }

        public static string ConsolidateWhitespace(string s)
        {
            return Regex.Replace(s, @"\s+", " ");
        }


        static private void writeTableToFile(string outputPath, List<TableRow> table)
        {
            using (StreamWriter fsw = File.CreateText(outputPath + "_tbl.csv"))
            {
                foreach (var row in table)
                {
                    foreach (var cell in row.Cells)
                    {
                        fsw.Write(cell.Text + '\t');
                    }
                    fsw.WriteLine();

                }
            }
        }
    }

    public class TableCell
    {
        public string Text;
        public enum HORIZONTAL_ALIGNMENT { UNKNOWN, LEFT, CENTER, RIGHT };
        public HORIZONTAL_ALIGNMENT HorizontalAlignment;
        public double Indentation;      // Dimensionless.  May be px, may be pt.
        public int IndentationLevel;    // Relative to all indentations of this column
        public bool Bold;

        public TableCell(AngleSharp.Dom.IElement element = null)
        {
            Text = "";
            HorizontalAlignment = HORIZONTAL_ALIGNMENT.UNKNOWN;
            Indentation = 0.0;
            IndentationLevel = 0;
            Bold = false;

            if (element != null)
            {
                Initialize(element);
            }
        }

        private void Initialize(AngleSharp.Dom.IElement element)
        {

            if (Text.Length == 0)       // Only take text from the first element that offers it.
            {
                // Replace an <br>'s with spaces.
                while (element.InnerHtml.Contains("<br>") || element.InnerHtml.Contains("<BR>"))
                {
                    element.InnerHtml = element.InnerHtml.Replace("<br>", " ").Replace("<BR>", " ");
                }

                Text = element.TextContent.Trim().Replace('\u00A0', ' ');       // Transform nbsp to regular space...
            }

            // Take other attributes from deepest last-sibling element that offers it.
            if (element.Style.TextAlign.ToLower() == "left")
            {
                HorizontalAlignment = HORIZONTAL_ALIGNMENT.LEFT;
            }
            else if (element.Style.TextAlign.ToLower() == "center")
            {
                HorizontalAlignment = HORIZONTAL_ALIGNMENT.CENTER;
            }
            else if (element.Style.TextAlign.ToLower() == "right")
            {
                HorizontalAlignment = HORIZONTAL_ALIGNMENT.RIGHT;
            }

            if (element.Style.FontWeight.ToLower() == "bold")
            {
                Bold = true;
            }

            // Questionable assumption: html authors don't mix px and pt dimensions in the same page...  
            if (element.Style.PaddingLeft.Length > 0)
            {
                double leftPad = 0;
                Double.TryParse(element.Style.PaddingLeft.Replace("px", "").Replace("pt", ""), out leftPad);
                this.Indentation += leftPad;
            }
            if (element.Style.MarginLeft.Length > 0)
            {
                double leftMargin = 0;
                Double.TryParse(element.Style.MarginLeft.Replace("px", "").Replace("pt", ""), out leftMargin);
                this.Indentation += leftMargin;
            }
            if (element.Style.TextIndent.Length > 0)
            {
                double textIndent = 0;
                Double.TryParse(element.Style.TextIndent.Replace("px", "").Replace("pt", ""), out textIndent);
                this.Indentation += textIndent;
            }

            // Eliminate "meaningless" indentation diffs.  MAKE THE SENSITIVITY OF THIS OVERRIDEABLE THROUGH CONFIG PARAMS.
            const int INDENTATION_GRANULARITY = 2;
            this.Indentation = (double)((((int)this.Indentation) / INDENTATION_GRANULARITY) * INDENTATION_GRANULARITY);

            // Iterate over this element's children, recursing down the branches of each.
            //  Recursion bounded by leaf elements having no children.
            foreach (var childElement in element.Children)
            {
                this.Initialize(childElement);
            }
        }

        public string CleanedNumericText
        {
            get
            {
                // Change wrapping parens to leading minus sign, drop any commas
                return Text.Replace('(', '-').Replace(")", "").Replace(",", "");
            }
        }

        public bool IsNumeric
        {
            get
            {
                return Regex.IsMatch(CleanedNumericText, @"\(?\d+\)?");
            }
        }
    }

    public class TableRow
    {
        public TableRow()
        {
            RowHead = null;
            Cells = new List<TableCell>();
            bRowCellsHaveContent = false;
            parentRow = null;
        }

        public void AddCell(TableCell cell)
        {
            // First TableCell added to a row is the RowHead
            if (RowHead == null)
            {
                RowHead = cell;
                RowHead.Text = RowHead.Text.Trim();
                if (RowHead.Text.EndsWith(":"))
                {
                    RowHead.Text = RowHead.Text.Substring(0, RowHead.Text.Length - 1);
                }
            }
            else
            {
                Cells.Add(cell);

                if (cell.Text.Length > 0)
                {
                    bRowCellsHaveContent = true;
                }
            }

        }

        public TableCell RowHead;
        public List<TableCell> Cells;
        public bool bRowCellsHaveContent;
        public TableRow parentRow;

    }


    class FlattenedRow
    {
        public FlattenedRow(string _attribute, string _time, string _value)
        {
            attribute = _attribute;
            time = _time;
            value = _value;
        }

        public string attribute;
        public string time;
        public string value;

        public override string ToString()
        {
            return attribute + "\t" + time + "\t" + value;
        }

        public override bool Equals(object obj)
        {
            FlattenedRow other = (FlattenedRow)obj;
            return (attribute.Equals(other.attribute) && time.Equals(other.time) && value.Equals(other.value));
        }

        public override int GetHashCode()
        {
            return new { attribute, time, value }.GetHashCode();
        }
    }
}
