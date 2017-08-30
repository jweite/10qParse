﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AngleSharp;
using System.Text.RegularExpressions;

namespace AngleSharpTest1
{
    class Program
    {
        static void Main(string[] args)
        {
            ExtractTableFromHTML(@"c:\temp\de-20170430x10q.htm", "STATEMENT OF CONSOLIDATED INCOME", 1, @"c:\temp\deere_consolidated_income_3mon", new Dictionary<string, string> { { "UndifferentiatedTotalAssociatesWithPrecedingBoldHeading", "true" } });
            ExtractTableFromHTML(@"c:\temp\de-20170430x10q.htm", "STATEMENT OF CONSOLIDATED COMPREHENSIVE INCOME", 1, @"c:\temp\deere_consolidated_comprehensive_income_3mon", new Dictionary<string, string> { { "UndifferentiatedTotalAssociatesWithPrecedingBoldHeading", "true" } });
            ExtractTableFromHTML(@"c:\temp\de-20170430x10q.htm", "STATEMENT OF CONSOLIDATED INCOME", 2, @"c:\temp\deere_consolidated_income_6mon", new Dictionary<string, string> { { "UndifferentiatedTotalAssociatesWithPrecedingBoldHeading", "true" } });
            ExtractTableFromHTML(@"c:\temp\de-20170430x10q.htm", "STATEMENT OF CONSOLIDATED COMPREHENSIVE INCOME", 2, @"c:\temp\deere_consolidated_comprehensive_income_6mon", new Dictionary<string, string> { { "UndifferentiatedTotalAssociatesWithPrecedingBoldHeading", "true" } });
            ExtractTableFromHTML(@"c:\temp\de-20170430x10q.htm", "CONDENSED CONSOLIDATED BALANCE SHEET", 1, @"c:\temp\deere_balance", null);
            ExtractTableFromHTML(@"c:\temp\de-20170430x10q.htm", "STATEMENT OF CONSOLIDATED CASH FLOWS", 1, @"c:\temp\deere_cashflow_6mon", new Dictionary<string, string> { { "UndifferentiatedTotalAssociatesWithPrecedingBoldHeading", "true" } });

            ExtractTableFromHTML(@"c:\temp\cat_10qx6302017.htm", "Consolidated Statement of Results of Operations", 1, @"c:\temp\cat", null);
            ExtractTableFromHTML(@"c:\temp\a17-13367_110q.htm", "CONSOLIDATED STATEMENT OF EARNINGS", 1, @"c:\temp\ibm", null);
            ExtractTableFromHTML(@"c:\temp\a10-qq32017712017.htm", "CONDENSED CONSOLIDATED STATEMENTS OF OPERATIONS (Unaudited)", 1, @"c:\temp\apple", null);
            ExtractTableFromHTML(@"c:\temp\mdt-2015q3x10q.htm", "CONDENSED CONSOLIDATED STATEMENTS OF EARNINGS", 1, @"c:\temp\medtronic", null);
            ExtractTableFromHTML(@"c:\temp\amzn-20170630x10q.htm", "CONSOLIDATED STATEMENTS OF OPERATIONS", 1, @"c:\temp\amazon", null);

            Console.Write("Done. Press enter to exit");
            Console.ReadLine();
        }


        public static void ExtractTableFromHTML(string sourcePath, string landmark, int landmarkIndex, string outputPath, IDictionary<string, string> config)
        {
            // Set up angleSharp html parser
            var angleSharConfig = Configuration.Default.WithCss();
            var parser = new AngleSharp.Parser.Html.HtmlParser(angleSharConfig);

            // Read the source html file
            FileStream fs = File.OpenRead(sourcePath);

            // Parse it into a DOM with AngleSharp
            var doc = parser.Parse(fs);

            // Select all the DOM's elements
            var tableSelector = doc.QuerySelectorAll("*");

            // Look for the specified occurence (landmarkIndex), and find the table related to it.  
            //  The related table is either the table containing the landmark (if the landmark is contained in a table) 
            //  or the table immeiately following the landmark (if the landmark isn't contained in a table)
            //  The below is a bit of a state machine, with the bool foundLandmark and in landmarkIndex controlling how it behaves
            //  as it scans the elements sequentially.

            bool foundLandmark = false;
            AngleSharp.Dom.IElement foundTable = null;
            for (int iElement = 0; iElement < tableSelector.Length; ++iElement)
            {
                var element = tableSelector[iElement];

                if (foundLandmark == false && element.TextContent.Trim() == landmark)
                {
                    if (--landmarkIndex > 0)           // Find the nth instance of the landmark...
                    {
                        // Skip past immediately subsequent subsequent elements with the same text (ie nested tags) before continuing the search
                        for (++iElement;  iElement < tableSelector.Length && tableSelector[++iElement].TextContent.Trim() == landmark; ++iElement)
                            ;
                        continue;
                    }

                    Console.WriteLine("Landmark found in tag " + element.TagName);

                    foundLandmark = true;

                    // If landmark is in a table, then that table is what we want.  Otherwise the next table is the one we want.
                    var element2 = element;
                    while (element2 != null && element2.TagName != "TABLE")
                    {
                        element2 = element2.ParentElement;
                    }
                    if (element2 != null && element2.TagName == "TABLE")
                    {
                        Console.WriteLine("Table found, containing landmark");
                        foundTable = element2;
                        break;
                    }
                }
                if (foundLandmark == true && element.TagName == "TABLE")
                {
                    Console.WriteLine("Table found below landmark.");
                    foundTable = element;
                    break;
                }
            }
            if (foundTable == null)
            {
                Console.WriteLine("No landmarked table found");
                return;
            }

            // Parse HTML table just found into a 2d matrix (actually a list of TableRow objects, which contain a list of TableCell objects)
            var tableData = new List<TableRow>();
            var rowElements = foundTable.QuerySelectorAll("TR");
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

            // Save a csv of the Matrix for analysis
//            writeTableToFile(outputPath, tableData);

            // Extract Headings
            IList<string> columnHeadings = ExtractColumnHeadings(tableData);
            if (columnHeadings == null)
            {
                Console.WriteLine("FATAL: Cannot find any qualifying heading rows in table");
            }

            CalcRowheadIndentationLevels(tableData);        // Calculate the relative indentation level of each rowhead vs the rest.

            BuildComplexRowHeads(tableData, config);


            // Flatten matrix with some rules
            List <FlattenedRow> results = new List<FlattenedRow>();
            string attributeName = "";

            for (int iRow = 0; iRow < tableData.Count; ++iRow) {
                var row = tableData[iRow];

                int nCols = row.Cells.Count;

                attributeName = row.RowHead.Text;

                TableRow row2 = row.parentRow;
                while (row2 != null && row2.RowHead.Text.Length > 0)
                {
                    attributeName = row2.RowHead.Text + "|" + attributeName;
                    row2 = row2.parentRow;
                }

                if (attributeName.Length == 0) continue;    // Assumption: rows without attribute names should be skipped.

                // Standardize attribute: only single spaces between words
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

                    // Fix up chars in col content
                    colContent = colContent.Replace('(', '-').Replace(")", "").Replace(",", "");     // Parens = -, drop commas.

                    // Get the heading for this col.  ASSUMTPION: Column headings are first contiguous cells with centered text.
                    string heading = columnHeadings[iCol];

                    FlattenedRow flatRow = new FlattenedRow(attributeName, heading, colContent);
                    results.Add(flatRow);

                }

            }

            // The spans result in duplicated entries.  Get rid of them.
            //  I'm using this strategy because the numeric value may or may not appear in some or all of the spanned cols.
            //  Easier to output all that it appears in, and then eliminate the dupes.
        
            var distinctFlatRows = results.Distinct();

            // Write the flattened matrix out
            using (StreamWriter fsw = File.CreateText(outputPath + ".csv"))
            {
                foreach (var record in distinctFlatRows)
                {
                    fsw.Write(record + "\r\n");
                }
            }
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

        public static void BuildComplexRowHeads(List<TableRow> tableData, IDictionary<string, string> config)
        {
            for (int iRow = 0; iRow < tableData.Count; ++iRow)
            {
                var row = tableData[iRow];

                // If this rowhead is one of a handful of values then it is a stand-alone heading (has no parent, even if indented)
                // MOVE THIS TO CONFIG, PROBABLY ON A PER COMPANY/FORM BASIS
                string[] standAloneRowheads = {
                    "Gross margin",      // Apple
                    "Total net sales", "Total non-operating income (expense)"    // Amazon
                };
                bool foundMatch = false;
                foreach (var standAloneRowhead in standAloneRowheads)
                {
                    if (row.RowHead.Text.ToLower() == standAloneRowhead.ToLower())
                    {
                        row.parentRow = null;
                        foundMatch = true;
                        break;
                    }
                }
                if (foundMatch == true)
                {
                    continue;
                }

                // If configured to associate undifferentiated "Total" headings with the preceding bold heading (ie Deere Income)
                foundMatch = false;
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

                // If this rowhead is of the form "Total XXXX" and XXX is a preceeding rowhead then this is a stand-alone heading (has no parent, even if indented).
                string s = row.RowHead.Text.ToLower();
                Match m = Regex.Match(s, @"total (.*)");
                foundMatch = false;
                if (m.Success)
                {
                    string lookFor = m.Groups[1].Captures[0].Value;
                    for (int iRow2 = iRow-1; iRow2 >= 0 && tableData[iRow2].RowHead.Text.Length > 0; --iRow2)
                    {
                        var precedingRow = tableData[iRow2];
                        if (precedingRow.RowHead.Text.ToLower() == lookFor)
                        {
                            row.parentRow = null;
                            foundMatch = true;
                            break;
                        }
                    }
                }
                if (foundMatch == true)
                {
                    continue;
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
                else if (!row.RowHead.Bold && row.bRowCellsHaveContent)
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
