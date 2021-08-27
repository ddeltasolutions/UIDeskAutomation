using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// This class represents a Data Grid control.
    /// </summary>
    public class UIDA_DataGrid: DataGridBase //ElementBase
    {
        /// <summary>
        /// Creates a DataGrid using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_DataGrid(AutomationElement el): base(el)
        {
            //this.uiElement = el;
        }

        /// <summary>
        /// Gets the header of a datagrid control.
        /// </summary>
        public DataGridHeader Header
        {
            get
            {
                AutomationElement header = this.FindFirst(ControlType.Header, null,
                    false, false, true);

                if (header == null)
                {
                    Engine.TraceInLogFile("DataGrid Header not found");

                    if (Engine.ThrowExceptionsWhenSearch == true)
                    {
                        throw new Exception("DataGrid Header not found");
                    }
                    else
                    {
                        return null;
                    }
                }

                DataGridHeader dataGridHeader = new DataGridHeader(header);
                return dataGridHeader;
            }
        }

        /// <summary>
        /// This class represents a header in datagrid control.
        /// </summary>
        public class DataGridHeader: DataGridBase
        {
            /// <summary>
            /// Creates a DataGridHeader using an AutomationElement
            /// </summary>
            /// <param name="el">UI Automation AutomationElement</param>
            public DataGridHeader(AutomationElement el): base(el)
            {
                //this.uiElement = el;
            }

            /// <summary>
            /// Header items in a datagrid header.
            /// </summary>
            public UIDA_HeaderItem[] Items
            {
                get
                {
                    List<AutomationElement> items = this.FindAll(ControlType.HeaderItem,
                        null, false, false, true);

                    List<UIDA_HeaderItem> headerItems = new List<UIDA_HeaderItem>();

                    foreach (AutomationElement item in items)
                    { 
                        UIDA_HeaderItem headerItem = new UIDA_HeaderItem(item);
                        headerItems.Add(headerItem);
                    }

                    return headerItems.ToArray();
                }
            }
        }

        /// <summary>
        /// Gets the column count.
        /// </summary>
        public int ColumnCount
        {
            get
            {
                GridPattern gridPattern = this.GetGridPattern();

                if (gridPattern != null)
                {
                    try
                    {
                        return gridPattern.Current.ColumnCount;
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile(
                            "DataGrid.ColumnCount - cannot get column count");

                        throw new Exception(
                            "DataGrid.ColumnCount - cannot get column count");
                    }
                }

                Engine.TraceInLogFile("DataGrid.ColumnCount - GridPattern not supported");
                throw new Exception("DataGrid.ColumnCount - GridPattern not supported");
            }
        }

        /// <summary>
        /// Gets the number of rows.
        /// </summary>
        public int RowCount
        {
            get
            {
                GridPattern gridPattern = this.GetGridPattern();

                if (gridPattern != null)
                {
                    try
                    {
                        return gridPattern.Current.RowCount;
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile("DataGrid.RowCount - cannot get row count");
                        throw new Exception("DataGrid.RowCount - cannot get row count");
                    }
                }

                Engine.TraceInLogFile("DataGrid.RowCount - GridPattern not supported");
                throw new Exception("DataGrid.RowCount - GridPattern not supported");
            }
        }

        /// <summary>
        /// Gets a boolean that tells if multiple items can be selected in the data grid.
        /// </summary>
        public bool CanSelectMultiple
        {
            get
            {
                SelectionPattern selectionPattern = this.GetSelectionPattern();

                if (selectionPattern != null)
                {
                    try
                    {
                        return selectionPattern.Current.CanSelectMultiple;
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile("DataGrid.CanSelectMultiple failed: " +
                            ex.Message);

                        throw new Exception("DataGrid.CanSelectMultiple failed: " +
                            ex.Message);
                    }
                }
                else
                {
                    Engine.TraceInLogFile(
                        "DataGrid.CanSelectMultiple - SelectionPattern not supported");

                    throw new Exception(
                        "DataGrid.CanSelectMultiple - SelectionPattern not supported");
                }
            }
        }

        /// <summary>
        /// Returns all rows in the current DataGrid control.
        /// </summary>
        public UIDA_DataItem[] Rows
        {
            get
            {
                if (uiElement.Current.FrameworkId != "WPF")
                {
                    List<AutomationElement> rows = this.FindAll(ControlType.DataItem,
                        null, true, false, true);

                    List<UIDA_DataItem> returnRows = new List<UIDA_DataItem>();

                    foreach (AutomationElement row in rows)
                    {
                        UIDA_DataItem dataItem = new UIDA_DataItem(row);
                        returnRows.Add(dataItem);
                    }

                    return returnRows.ToArray();
                }
                else
                {
                    List<UIDA_DataItem> returnRows = new List<UIDA_DataItem>();
                    
                    object objectPattern = null;
                    this.uiElement.TryGetCurrentPattern(ItemContainerPattern.Pattern, out objectPattern);
                    ItemContainerPattern itemContainerPattern = objectPattern as ItemContainerPattern;
                    if (itemContainerPattern == null)
                    {
                        List<AutomationElement> rows = this.FindAll(ControlType.DataItem,
                            null, true, false, true);
                        foreach (AutomationElement row in rows)
                        {
                            UIDA_DataItem dataItem = new UIDA_DataItem(row);
                            returnRows.Add(dataItem);
                        }
                        return returnRows.ToArray();
                    }
                    
                    AutomationElement crt = null;
                    do
                    {
                        crt = itemContainerPattern.FindItemByProperty(crt, null, null);
                        if (crt != null)
                        {
                            UIDA_DataItem dataItem = new UIDA_DataItem(crt);
                            returnRows.Add(dataItem);
                        }
                    }
                    while (crt != null);
                    
                    return returnRows.ToArray();
                }
            }
        }
        
        /// <summary>
        /// Gets the row at the specified index.
        /// </summary>
        /// <param name="index">zero based index</param>
        public UIDA_DataItem this[int index]
        {
            get
            {
                if (uiElement.Current.FrameworkId != "WPF")
                {
                    return this.Rows[index];
                }
                else
                {
                    object objectPattern = null;
                    this.uiElement.TryGetCurrentPattern(ItemContainerPattern.Pattern, out objectPattern);
                    ItemContainerPattern itemContainerPattern = objectPattern as ItemContainerPattern;
                    
                    if (itemContainerPattern == null)
                    {
                        return Rows[index];
                    }
                    else
                    {
                        if (index < 0)
                        {
                            throw new Exception("Index cannot be negative");
                        }
                        AutomationElement crt = null;
                        do
                        {
                            crt = itemContainerPattern.FindItemByProperty(crt, null, null);
                            if (crt == null)
                            {
                                throw new Exception("Index too big");
                            }
                            index--;
                        }
                        while (index >= 0);
                        
                        UIDA_DataItem dataItem = new UIDA_DataItem(crt);
                        return dataItem;
                    }
                }
            }
        }

        /// <summary>
        /// Gets all groups in a DataGrid control.
        /// </summary>
        new public DataGridGroup[] Groups
        {
            get
            {
                List<AutomationElement> allGroups = this.FindAll(ControlType.Group,
                    null, false, false, true);

                List<DataGridGroup> returnGroups = new List<DataGridGroup>();

                foreach (AutomationElement group in allGroups)
                {
                    DataGridGroup returnGroup = new DataGridGroup(group);
                    returnGroups.Add(returnGroup);
                }

                return returnGroups.ToArray();
            }
        }

        /// <summary>
        /// Selects all items in a DataGrid control.
        /// </summary>
        public void SelectAll()
        {
            UIDA_DataItem[] items = this.Rows;

            foreach (UIDA_DataItem item in items)
            {
                try
                {
                    item.AddToSelection();
                }
                catch (Exception ex)
                { }
            }
        }

        /// <summary>
        /// Clears all selection in a DataGrid control.
        /// </summary>
        public void ClearAllSelection()
        {
            UIDA_DataItem[] selectedItems = this.SelectedRows;

            foreach (UIDA_DataItem selectedItem in selectedItems)
            {
                try
                {
                    selectedItem.RemoveFromSelection();
                }
                catch (Exception ex)
                { }
            }
        }

        /// <summary>
        /// Returns selected rows in the current DataGrid control.
        /// </summary>
        public UIDA_DataItem[] SelectedRows
        {
            get
            {
                SelectionPattern selectionPattern = this.GetSelectionPattern();

                if (selectionPattern == null)
                {
                    Engine.TraceInLogFile(
                        "DataGrid.SelectedRows - SelectionPattern not supported");

                    throw new Exception(
                        "DataGrid.SelectedRows - SelectionPattern not supported");
                }

                try
                {
                    AutomationElement[] selectedItems =
                        selectionPattern.Current.GetSelection();

                    List<UIDA_DataItem> returnCollection = new List<UIDA_DataItem>();

                    foreach (AutomationElement selectedItem in selectedItems)
                    {
                        UIDA_DataItem dataItem = new UIDA_DataItem(selectedItem);
                        returnCollection.Add(dataItem);
                    }

                    return returnCollection.ToArray();
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile(
                        "DataGrid.SelectedRows failed: " + ex.Message);

                    throw new Exception(
                        "DataGrid.SelectedRows failed: " + ex.Message);
                }
            }
        }
        
        /// <summary>
        /// Scrolls the DataGrid vertically and horizontally using the specified percents.
        /// </summary>
        /// <param name="percentVertical">percentage to scroll vertically</param>
        /// <param name="percentHorizontal">percentage to scroll horizontally</param>
        public void Scroll(double percentVertical, double percentHorizontal = 0)
        {
            object objectPattern = null;
            this.uiElement.TryGetCurrentPattern(ScrollPattern.Pattern, out objectPattern);
            ScrollPattern scrollPattern = objectPattern as ScrollPattern;
            
            if (scrollPattern != null)
            {
                try
                {
                    scrollPattern.SetScrollPercent(percentHorizontal, percentVertical);
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("Scroll(): " + ex.Message);
                }
            }
        }
        
        private AutomationElement GetItemByText(string text)
        {
            object objectPattern = null;
            this.uiElement.TryGetCurrentPattern(ItemContainerPattern.Pattern, out objectPattern);
            ItemContainerPattern itemContainerPattern = objectPattern as ItemContainerPattern;
            
            if (itemContainerPattern != null)
            {
                return itemContainerPattern.FindItemByProperty(null, AutomationElement.NameProperty, text);
            }
            return null;
        }
        
        private AutomationElement GetWPFDataItem(int index)
        {
            object objectPattern = null;
            this.uiElement.TryGetCurrentPattern(ItemContainerPattern.Pattern, out objectPattern);
            ItemContainerPattern itemContainerPattern = objectPattern as ItemContainerPattern;
            
            if (itemContainerPattern == null)
            {
                return null;
            }
            
            AutomationElement crt = null;
            do
            {
                crt = itemContainerPattern.FindItemByProperty(crt, null, null);
                if (crt == null)
                {
                    Engine.TraceInLogFile("list item index too big");
                    break;
                }
                index--;
            }
            while (index != 0);
            
            if (index == 0)
            {
                return crt;
            }
            else
            {
                return null;
            }
        }
        
        /// <summary>
        /// Selects an item in a DataGrid by the item index. Other selected items will be deselected.
        /// </summary>
        /// <param name="index">item index, starts with 1</param>
        public void Select(int index)
        {
            if (uiElement.Current.FrameworkId == "WPF")
            {
                AutomationElement item = GetWPFDataItem(index);
                if (item != null)
                {
                    UIDA_DataItem dataItemWPF = new UIDA_DataItem(item);
                    dataItemWPF.Select();
                    return;
                }
            }
            
            UIDA_DataItem dataItem = DataItemAt(null, index, true);
            if (dataItem == null)
            {
                Engine.TraceInLogFile("Item not found");
                throw new Exception("Item not found");
            }
            dataItem.Select();
        }
        
        /// <summary>
        /// Adds an item to selection in a DataGrid by the item index.
        /// </summary>
        /// <param name="index">item index, starts with 1</param>
        public void AddToSelection(int index)
        {
            if (uiElement.Current.FrameworkId == "WPF")
            {
                AutomationElement item = GetWPFDataItem(index);
                if (item != null)
                {
                    UIDA_DataItem dataItemWPF = new UIDA_DataItem(item);
                    dataItemWPF.AddToSelection();
                    return;
                }
            }
            
            UIDA_DataItem dataItem = DataItemAt(null, index, true);
            if (dataItem == null)
            {
                Engine.TraceInLogFile("Item not found");
                throw new Exception("Item not found");
            }
            dataItem.AddToSelection();
        }
        
        /// <summary>
        /// Removes an item from selection in a DataGrid by the item index.
        /// </summary>
        /// <param name="index">item index, starts with 1</param>
        public void RemoveFromSelection(int index)
        {
            if (uiElement.Current.FrameworkId == "WPF")
            {
                AutomationElement item = GetWPFDataItem(index);
                if (item != null)
                {
                    UIDA_DataItem dataItemWPF = new UIDA_DataItem(item);
                    dataItemWPF.RemoveFromSelection();
                    return;
                }
            }
            
            UIDA_DataItem dataItem = DataItemAt(null, index, true);
            if (dataItem == null)
            {
                Engine.TraceInLogFile("Item not found");
                throw new Exception("Item not found");
            }
            dataItem.RemoveFromSelection();
        }

        private SelectionPattern GetSelectionPattern()
        {
            object selectionPatternObj = null;

            this.uiElement.TryGetCurrentPattern(SelectionPattern.Pattern,
                out selectionPatternObj);

            SelectionPattern selectionPattern = selectionPatternObj as
                SelectionPattern;

            return selectionPattern;
        }

        private GridPattern GetGridPattern()
        {
            object gridPatternObj = null;

            if (this.uiElement.TryGetCurrentPattern(GridPattern.Pattern,
                out gridPatternObj) == true)
            {
                GridPattern gridPattern = gridPatternObj as GridPattern;

                return gridPattern;
            }

            return null;
        }
    }

    /// <summary>
    /// This class represents a header item in a datagrid.
    /// </summary>
    public class UIDA_HeaderItem : ElementBase
    {
        /// <summary>
        /// Creates a UDA_HeaderItem using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_HeaderItem(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Text of header item.
        /// </summary>
        public string Text
        {
            get
            {
                string textString = null;

                try
                {
                    textString = uiElement.Current.Name;
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("HeaderItem text: " + ex.Message);
                    throw new Exception("HeaderItem text: " + ex.Message);
                }

                return textString;
            }
        }
    }

    /// <summary>
    /// This class represents a group.
    /// </summary>
    public class DataGridGroup : DataGridBase
    {
        /// <summary>
        /// Creates a DataGridGroup using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public DataGridGroup(AutomationElement el)
            : base(el)
        { }

        /// <summary>
        /// Expands a group inside a data grid.
        /// </summary>
        public void Expand()
        {
            ExpandCollapsePattern expandCollapsePattern = this.GetExpandCollapsePattern();

            if (expandCollapsePattern == null)
            {
                Engine.TraceInLogFile(
                    "DataGridGroup.Expand() - ExpandCollapsePattern not supported");

                throw new Exception(
                    "DataGridGroup.Expand() - ExpandCollapsePattern not supported");
            }

            try
            {
                if (expandCollapsePattern.Current.ExpandCollapseState !=
                    ExpandCollapseState.Expanded)
                {
                    expandCollapsePattern.Expand();
                }
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("DataGridGroup.Expand() error: " + ex.Message);
                throw new Exception("DataGridGroup.Expand() error: " + ex.Message);
            }
        }

        /// <summary>
        /// Collapses a group inside a data grid.
        /// </summary>
        public void Collapse()
        {
            ExpandCollapsePattern expandCollapsePattern =
                this.GetExpandCollapsePattern();

            if (expandCollapsePattern == null)
            {
                Engine.TraceInLogFile(
                    "DataGridGroup.Collapse() - ExpandCollapsePattern not supported");

                throw new Exception(
                    "DataGridGroup.Collapse() - ExpandCollapsePattern not supported");
            }

            try
            {
                if (expandCollapsePattern.Current.ExpandCollapseState !=
                    ExpandCollapseState.Collapsed)
                {
                    expandCollapsePattern.Collapse();
                }
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("DataGridGroup.Collapse() error: " + ex.Message);
                throw new Exception("DataGridGroup.Collapse() error: " + ex.Message);
            }
        }

        private ExpandCollapsePattern GetExpandCollapsePattern()
        {
            object expandCollapsePatternObj = null;

            if (this.uiElement.TryGetCurrentPattern(ExpandCollapsePattern.Pattern,
                out expandCollapsePatternObj) == true)
            {
                ExpandCollapsePattern expandCollapsePattern =
                    expandCollapsePatternObj as ExpandCollapsePattern;

                return expandCollapsePattern;
            }

            return null;
        }
    }

    /// <summary>
    /// This class represents a data item.
    /// </summary>
    public class UIDA_DataItem : DataGridBase
    {
        /// <summary>
        /// Creates a UIDA_DataItem using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_DataItem(AutomationElement el)
            : base(el)
        { }

        /// <summary>
        /// Gets/sets the selection state of a DataItem in a DataGrid.
        /// </summary>
        public bool IsSelected
        {
            get
            {
                SelectionItemPattern selectionItemPattern = this.GetSelectionItemPattern();

                if (selectionItemPattern == null)
                {
                    Engine.TraceInLogFile(
                        "DataItem.IsSelected - SelectionItemPattern not supported");

                    throw new Exception(
                        "DataItem.IsSelected - SelectionItemPattern not supported");
                }

                try
                {
                    return selectionItemPattern.Current.IsSelected;
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile(
                        "DataItem.IsSelected failed: " + ex.Message);

                    throw new Exception(
                        "DataItem.IsSelected failed: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Select the current item and deselects all other selected items.
        /// </summary>
        public void Select()
        {
            SelectionItemPattern selectionItemPattern = this.GetSelectionItemPattern();

            if (selectionItemPattern == null)
            {
                Engine.TraceInLogFile(
                    "DataItem.Select() - SelectionItemPattern not supported");

                throw new Exception(
                    "DataItem.Select() - SelectionItemPattern not supported");
            }

            try
            {
                selectionItemPattern.Select();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("DataItem.Select() failed: " + ex.Message);
                throw new Exception("DataItem.Select() failed: " + ex.Message);
            }
        }

        /// <summary>
        /// Adds the current DataItem to the selected items.
        /// </summary>
        public void AddToSelection()
        {
            SelectionItemPattern selectionItemPattern = this.GetSelectionItemPattern();

            if (selectionItemPattern == null)
            {
                Engine.TraceInLogFile(
                    "DataItem.AddToSelection() - SelectionItemPattern not supported");

                throw new Exception(
                    "DataItem.AddToSelection() - SelectionItemPattern not supported");
            }

            try
            {
                selectionItemPattern.AddToSelection();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile(
                    "DataItem.AddToSelection() failed: " + ex.Message);

                throw new Exception(
                    "DataItem.AddToSelection() failed: " + ex.Message);
            }
        }

        /// <summary>
        /// Removes the current DataItem from selected items.
        /// </summary>
        public void RemoveFromSelection()
        {
            SelectionItemPattern selectionItemPattern = this.GetSelectionItemPattern();

            if (selectionItemPattern == null)
            {
                Engine.TraceInLogFile(
                    "DataItem.RemoveFromSelection() - SelectionItemPattern not supported");

                throw new Exception(
                    "DataItem.RemoveFromSelection() - SelectionItemPattern not supported");
            }

            try
            {
                selectionItemPattern.RemoveFromSelection();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile(
                    "DataItem.RemoveFromSelection() failed: " + ex.Message);

                throw new Exception(
                    "DataItem.RemoveFromSelection() failed: " + ex.Message);
            }
        }

        private SelectionItemPattern GetSelectionItemPattern()
        {
            object selectionItemPatternObj = null;

            this.uiElement.TryGetCurrentPattern(SelectionItemPattern.Pattern, 
                out selectionItemPatternObj);

            SelectionItemPattern selectionItemPattern =
                selectionItemPatternObj as SelectionItemPattern;

            return selectionItemPattern;
        }
        
        /// <summary>
        /// Gets the value at the specified column index.
        /// </summary>
        /// <param name="columnIndex">zero based column index</param>
        public string this[int columnIndex]
        {
            get
            {
                object objectPattern = null;
                this.uiElement.TryGetCurrentPattern(ItemContainerPattern.Pattern, out objectPattern);
                ItemContainerPattern itemContainerPattern = objectPattern as ItemContainerPattern;
                
                if (itemContainerPattern == null)
                {
                    AutomationElementCollection collection = uiElement.FindAll(TreeScope.Children, Condition.TrueCondition);
                    return (new UIDA_Custom(collection[columnIndex])).GetText();
                }
                else
                {
                    if (columnIndex < 0)
                    {
                        throw new Exception("Index cannot be negative");
                    }
                    AutomationElement crt = null;
                    do
                    {
                        crt = itemContainerPattern.FindItemByProperty(crt, null, null);
                        if (crt == null)
                        {
                            throw new Exception("Index too big");
                        }
                        columnIndex--;
                    }
                    while (columnIndex >= 0);
                    
                    return (new UIDA_Custom(crt)).GetText();
                }
            }
        }
        
        /// <summary>
        /// Gets the value at the specified column.
        /// </summary>
        /// <param name="columnName">column name</param>
        public string this[string columnName]
        {
            get
            {
                TreeWalker tw = TreeWalker.ControlViewWalker;
                AutomationElement gridEl = tw.GetParent(this.uiElement);
                UIDA_DataGrid grid = new UIDA_DataGrid(gridEl);
                
                UIDA_HeaderItem[] headerItems = grid.Header.Items;
                int columnIndex = -1;
                for (int i = 0; i < headerItems.Length; i++)
                {
                    if (columnName == headerItems[i].Text)
                    {
                        columnIndex = i;
                        break;
                    }
                }
                
                if (columnIndex >= 0)
                {
                    return this[columnIndex];
                }
                else
                {
                    throw new Exception("No column with this name");
                }
            }
        }
    }

    /// <summary>
    /// Represents base class for DataGrid control. This class cannot be instantiated.
    /// </summary>
    abstract public class DataGridBase: ElementBase
    {
        internal DataGridBase(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Searches a header item in the current element
        /// </summary>
        /// <param name="name">text of header item</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_HeaderItem element</returns>
        public UIDA_HeaderItem HeaderItem(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            AutomationElement returnElement = this.FindFirst(ControlType.HeaderItem,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("HeaderItem method - HeaderItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderItem method - HeaderItem element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_HeaderItem headerItem = new UIDA_HeaderItem(returnElement);
            return headerItem;
        }

        /// <summary>
        /// Searches for a HeaderItem with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of HeaderItem</param>
        /// <param name="index">index of HeaderItem</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_HeaderItem element</returns>
        public UIDA_HeaderItem HeaderItemAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("HeaderItemAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderItemAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            AutomationElement returnElement = null;

            Errors error = this.FindAt(ControlType.HeaderItem, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("HeadetItemAt method - HeaderItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderItemAt method - HeaderItem element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("HeaderItemAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("HeaderItemAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_HeaderItem headerItem = new UIDA_HeaderItem(returnElement);
            return headerItem;
        }

        /// <summary>
        /// Searches a group item in the current element
        /// </summary>
        /// <param name="name">text of group item</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>DataGridGroup element</returns>
        new public DataGridGroup Group(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            AutomationElement returnElement = this.FindFirst(ControlType.Group,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Group method - Group element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Group method - Group element not found");
                }
                else
                {
                    return null;
                }
            }

            DataGridGroup group = new DataGridGroup(returnElement);
            return group;
        }

        /// <summary>
        /// Searches for a Group item with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of group</param>
        /// <param name="index">index of group</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>DataGridGroup element</returns>
        new public DataGridGroup GroupAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("GroupAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GroupAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            AutomationElement returnElement = null;

            Errors error = this.FindAt(ControlType.Group, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("GroupAt method - Group element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GroupAt method - Group element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("GroupAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GroupAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            DataGridGroup group = new DataGridGroup(returnElement);
            return group;
        }

        /// <summary>
        /// Searches a data item in the current element
        /// </summary>
        /// <param name="name">text of data item</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_DataItem element</returns>
        public UIDA_DataItem DataItem(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            AutomationElement returnElement = this.FindFirst(ControlType.DataItem,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("DataItem method - DataItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataItem method - DataItem element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_DataItem dataItem = new UIDA_DataItem(returnElement);
            return dataItem;
        }

        /// <summary>
        /// Searches for a DataItem with a specified text at a specified index.
        /// </summary>
        /// <param name="name">text of DataItem</param>
        /// <param name="index">index of DataItem</param>
        /// <param name="searchDescendants">true if search through descendants, false if search only through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_DataItem element</returns>
        public UIDA_DataItem DataItemAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("DataItemAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataItemAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            AutomationElement returnElement = null;

            Errors error = this.FindAt(ControlType.DataItem, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("DataItemAt method - DataItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataItemAt method - DataItem element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("DataItemAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("DataItemAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_DataItem dataItem = new UIDA_DataItem(returnElement);
            return dataItem;
        }
        
        /// <summary>
        /// Returns a collection of DataItem that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of DataItem elements, use null to return all DataItems</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_DataItem elements</returns>
        public UIDA_DataItem[] DataItems(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<AutomationElement> allDataItems = FindAll(ControlType.DataItem,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_DataItem> dataitems = new List<UIDA_DataItem>();
            if (allDataItems != null)
            {
                foreach (AutomationElement crtEl in allDataItems)
                {
                    dataitems.Add(new UIDA_DataItem(crtEl));
                }
            }
            return dataitems.ToArray();
        }
    }
}
