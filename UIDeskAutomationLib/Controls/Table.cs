using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// This class represents a Table control.
    /// </summary>
    public class Table: ElementBase
    {
        public Table(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// This class represents a row in a table control.
        /// </summary>
        public class Row : ElementBase
        {
            public Row(AutomationElement el)
            {
                this.uiElement = el;
            }

            /// <summary>
            /// Gets a collection of headers for the current row.
            /// </summary>
            public Header[] Headers
            {
                get
                {
                    List<AutomationElement> headers = this.FindAll(ControlType.Header,
                        null, false, false, true);

                    List<Header> returnHeaders = new List<Header>();

                    foreach (AutomationElement header in headers)
                    {
                        Header returnHeader = new Header(header);
                        returnHeaders.Add(returnHeader);
                    }

                    return returnHeaders.ToArray();
                }
            }

            /// <summary>
            /// Searches a header in the current element
            /// </summary>
            /// <param name="name">text of header item</param>
            /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
            /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
            /// <returns>Header element</returns>
            public Header Header(string name = null, bool searchDescendants = false,
                bool caseSensitive = true)
            {
                AutomationElement returnElement = this.FindFirst(ControlType.Header,
                    name, searchDescendants, false, caseSensitive);

                if (returnElement == null)
                {
                    Engine.TraceInLogFile("Header method - Header element not found");

                    if (Engine.ThrowExceptionsWhenSearch == true)
                    {
                        throw new Exception("Header method - Header element not found");
                    }
                    else
                    {
                        return null;
                    }
                }

                Header header = new Header(returnElement);
                return header;
            }

            /// <summary>
            /// Searches for a Header with a specified text at a specified index.
            /// </summary>
            /// <param name="name">text of Header</param>
            /// <param name="index">index of Header</param>
            /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
            /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
            /// <returns>Header element</returns>
            public Header HeaderAt(string name, int index, bool searchDescendants = false,
                bool caseSensitive = true)
            {
                if (index < 0)
                {
                    Engine.TraceInLogFile("HeaderAt method - index cannot be negative");

                    if (Engine.ThrowExceptionsWhenSearch == true)
                    {
                        throw new Exception("HeaderAt method - index cannot be negative");
                    }
                    else
                    {
                        return null;
                    }
                }

                AutomationElement returnElement = null;

                Errors error = this.FindAt(ControlType.Header, name, index, searchDescendants,
                    false, caseSensitive, out returnElement);

                if (error == Errors.ElementNotFound)
                {
                    Engine.TraceInLogFile("HeaderAt method - Header element not found");

                    if (Engine.ThrowExceptionsWhenSearch == true)
                    {
                        throw new Exception("HeaderAt method - Header element not found");
                    }
                    else
                    {
                        return null;
                    }
                }
                else if (error == Errors.IndexTooBig)
                {
                    Engine.TraceInLogFile("HeaderAt method - index too big");

                    if (Engine.ThrowExceptionsWhenSearch == true)
                    {
                        throw new Exception("HeaderAt method - index too big");
                    }
                    else
                    {
                        return null;
                    }
                }

                Header header = new Header(returnElement);
                return header;
            }

            /// <summary>
            /// Gets a collection of Cells for the current row.
            /// </summary>
            public Cell[] Cells
            { 
                get
                {
                    List<AutomationElement> cells = this.FindAll(ControlType.Custom,
                            null, false, false, true);

                    List<Cell> returnCells = new List<Cell>();

                    foreach (AutomationElement cell in cells)
                    {
                        Cell returnCell = new Cell(cell);
                        returnCells.Add(returnCell);
                    }

                    return returnCells.ToArray();
                }
            }
        }

        /// <summary>
        /// This class represents a row header.
        /// </summary>
        public class Header : ElementBase
        {
            public Header(AutomationElement el)
            {
                this.uiElement = el;
            }
        }

        /// <summary>
        /// This class represents a cell in a row.
        /// </summary>
        public class Cell : ElementBase
        {
            public Cell(AutomationElement el)
            {
                this.uiElement = el;
            }

            /// <summary>
            /// Gets or sets the value of cell.
            /// </summary>
            public string Value
            {
                get
                {
                    return this.GetSetValue(null);
                }
                set
                {
                    GetSetValue(value, false);
                }
            }

            private string GetSetValue(string value, bool get = true)
            {
                object valuePatternObj = null;
                this.uiElement.TryGetCurrentPattern(ValuePattern.Pattern,
                    out valuePatternObj);

                ValuePattern valuePattern = valuePatternObj as ValuePattern;

                if (valuePattern == null)
                {
                    Engine.TraceInLogFile(
                        "Table.Cell get/set Value: ValuePattern not supported");

                    throw new Exception(
                        "Table.Cell get/set value: ValuePattern not supported");
                }

                if (get == true)
                {
                    string returnValue = string.Empty;

                    try
                    {
                        returnValue = valuePattern.Current.Value;
                        return returnValue;
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile(
                            "Table.Cell get Value failed: " + ex.Message);

                        throw new Exception(
                            "Table.Cell get Value failed: " + ex.Message);
                    }
                }
                else
                {
                    // Set value
                    try
                    {
                        valuePattern.SetValue(value);
                        return string.Empty;
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile(
                            "Table.Cell set Value failed: " + ex.Message);

                        throw new Exception(
                            "Table.Cell set Value failed: " + ex.Message);
                    }
                }
            }
        }

        /// <summary>
        /// Gets the top row of a table control.
        /// </summary>
        public Row TopRow
        {
            get
            {
                AutomationElement elFirstRow = this.FindFirst(ControlType.Custom, null,
                    false, false, true);

                if (elFirstRow == null)
                {
                    Engine.TraceInLogFile("Table.FirstRow - cannot find first row");

                    if (Engine.ThrowExceptionsWhenSearch == true)
                    {
                        throw new Exception("Table.FirstRow - cannot find first row");
                    }
                    else
                    {
                        return null;
                    }
                }

                Row topRow = new Row(elFirstRow);
                return topRow;
            }
        }

        /// <summary>
        /// Gets a collection with all rows of a table control.
        /// </summary>
        public Row[] Rows
        {
            get
            {
                List<AutomationElement> rows = this.FindAll(ControlType.Custom, null, 
                    false, false, true);

                List<Row> returnRows = new List<Row>();

                for (int i = 1; i < rows.Count; i++)
                {
                    Row row = new Row(rows[i]);
                    returnRows.Add(row);
                }

                return returnRows.ToArray();
            }
        }

        /// <summary>
        /// Gets the columns count of table control.
        /// </summary>
        public int ColumnsCount
        {
            get
            {
                return (this.TopRow.Headers.Length - 1);
            }
        }
    }
}
