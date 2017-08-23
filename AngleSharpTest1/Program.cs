using System;
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
            //ExtractTableFromHTML(@"c:\temp\cat_10qx6302017.htm", "Consolidated Statement of Results of Operations", @"c:\temp\cat");
            //ExtractTableFromHTML(@"c:\temp\a17-13367_110q.htm", "CONSOLIDATED STATEMENT OF EARNINGS", @"c:\temp\ibm");
            //ExtractTableFromHTML(@"c:\temp\a10-qq32017712017.htm", "CONDENSED CONSOLIDATED STATEMENTS OF OPERATIONS (Unaudited)", @"c:\temp\apple");
            ExtractTableFromHTML(@"c:\temp\de-20170430x10q.htm", "STATEMENT OF CONSOLIDATED INCOME", @"c:\temp\deere");
            //ExtractTableFromHTML(@"c:\temp\mdt-2015q3x10q.htm", "CONDENSED CONSOLIDATED STATEMENTS OF EARNINGS", @"c:\temp\medtronic");
            //ExtractTableFromHTML(@"c:\temp\amzn-20170630x10q.htm", "CONSOLIDATED STATEMENTS OF OPERATIONS", @"c:\temp\amazon");
            
        }


        public static void ExtractTableFromHTML(string sourcePath, string landmark, string outputPath)
        {
            var config = Configuration.Default.WithCss();
            var parser = new AngleSharp.Parser.Html.HtmlParser(config);

            FileStream fs = File.OpenRead(sourcePath);

            var doc = parser.Parse(fs);

            var tableSelector = doc.QuerySelectorAll("*");

            bool foundLandmark = false;
            AngleSharp.Dom.IElement foundTable = null;

            foreach(var element in tableSelector)
            {
                if (foundLandmark == false && element.TextContent.Trim() == landmark)
                {
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
                Console.WriteLine("Press enter to exit");
                Console.ReadLine();
                return;
            }

            // Parse HTML table into matrix (2d row-oriented list of lists)
            var tableData = new List<List<TableCell>>();
            var rowElements = foundTable.QuerySelectorAll("TR");
            foreach (var rowElement in rowElements)
            {
                var rowData = new List<TableCell>();

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
                        rowData.Add(tableCell);
                    }
                }

                tableData.Add(rowData);
            }

            // Do a little clean-up
            // dropCompletelyBlankRows(tableData);

            // Save a csv of the Matrix for analysis
            writeTableToFile(outputPath, tableData);

            // Flatten matrix it with some rules
            List<FlattenedRow> results = new List<FlattenedRow>();
            string attributeName = "";

            for (int iRow = 0; iRow < tableData.Count; ++iRow) {
                var row = tableData[iRow];

                int nCols = row.Count;

                // Assumption 1: attribute name is in column 0.
                attributeName = row[0].Text.Replace(":", "");
                if (attributeName.Length == 0) continue;    // Assumption: rows without attribute names should be skipped.

                // Special handling for rows labeled Basic or Diluted: find their overarching heading.
                //   Assumption:  Rows labeled only "Basic" and "Diluted" are futher identified by the heading immediately above them that's not "Basic" or "Diluted".
                if (attributeName == "Basic" || attributeName == "Diluted")
                {
                    
                    int iSuperHead = iRow - 1;
                    if (tableData[iSuperHead][0].Text == "Basic" || tableData[iSuperHead][0].Text == "Diluted")
                    {
                        --iSuperHead;
                    }
                    if (iSuperHead > 0)
                    {
                        attributeName = tableData[iSuperHead][0].Text + " " + attributeName;
                    }
                }

                // Special handling for rows labeled "Total": find their overarching heading.
                //  Assumption: a blank row will demarcate the sub-section of the table that this total is for.
                if (attributeName == "Total")
                {
                    int iSuperHead = 0;
                    for (iSuperHead = iRow; iSuperHead > 0 && tableData[iSuperHead][0].Text.Length > 0; --iSuperHead)
                        ;
                    if (tableData[iSuperHead][0].Text.Length == 0)
                    {
                        ++iSuperHead;
                    }
                    attributeName = attributeName + ": " + tableData[iSuperHead][0].Text;
                }


                // Scan columns
                string colContent = "";
                for (int iCol = 1; iCol < nCols; ++iCol)
                {
                    var col = row[iCol];
                    colContent = col.Text;

                    // Process numeric columns only.
                    if (!Regex.IsMatch(colContent, @"\(?\d+\)?"))
                    {
                        continue;
                    }

                    // Fix up chars in col content
                    colContent = colContent.Replace('(', '-').Replace(")", "").Replace(",", "");     // Parens = -, drop commas.

                    // Get the heading for this col.  ASSUMTPION: Column headings are first contiguous cells with centered text.
                    string heading = "";
                    bool centeredCellFound = false;
                    for (int ir = 0; ir < tableData.Count; ++ir)
                    {
                        TableCell cell = tableData[ir][iCol];
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
                        foreach(var row2 in tableData)
                        {
                            TableCell cell = row2[0];
                            string cellTextLower = cell.Text.ToLower();

                            // Assumption for now: these explanations are in column 0
                            Match mMonths = Regex.Match(cellTextLower, @"(\w*) months ended");
                            if (mMonths.Success)
                            {
                                string numMonths = mMonths.Groups[1].Captures[0].Value;
                                numMonths = numMonths.Substring(0, 1).ToUpper() + numMonths.Substring(1);

                                Match mDate = Regex.Match(cellTextLower, @"(\w+) (\d+)?, " + heading);
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

            Console.Write("Done. Press enter to exit");
            Console.ReadLine();
        }


        static private void writeTableToFile(string outputPath, List<List<TableCell>> table)
        {
            using (StreamWriter fsw = File.CreateText(outputPath + "_tbl.csv"))
            {
                foreach (var row in table)
                {
                    foreach (var cell in row)
                    {
                        fsw.Write(cell.Text + '\t');
                    }
                    fsw.WriteLine();

                }
            }
        }

        static private void dropCompletelyBlankRows(List<List<TableCell>> table)
        {
            for (int i = table.Count - 1; i >= 0; --i)
            {
                var row = table[i];
                bool rowHasContent = false;

                foreach (var col in row)
                {
                    if (col.Text.Length > 0)
                    {
                        rowHasContent = true;
                        break;
                    }
                }
                if (!rowHasContent)
                {
                    table.RemoveAt(i);
                }
            }
        }
    }

    class TableCell
    {
        public string Text;
        public enum HORIZONTAL_ALIGNMENT { UNKNOWN, LEFT, CENTER, RIGHT };
        public HORIZONTAL_ALIGNMENT HorizontalAlignment;
        public int Indentation;
        public bool Bold;

        public TableCell(AngleSharp.Dom.IElement element = null)
        {
            Text = "";
            HorizontalAlignment = HORIZONTAL_ALIGNMENT.UNKNOWN;
            Indentation = 0;
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

            if (element.Style.PaddingBottom.Length > 0)
            {
                int leftPad = 0;
                Int32.TryParse(element.Style.PaddingBottom.Replace("px", ""), out leftPad);
                if (leftPad > this.Indentation)
                {
                    this.Indentation = leftPad;
                }
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
