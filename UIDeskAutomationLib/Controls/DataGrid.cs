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
    public class UIDA_DataGrid: ElementBase
    {
        /// <summary>
        /// Creates a DataGrid using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_DataGrid(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets the header of a datagrid control.
        /// </summary>
        public UIDA_Header Header
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

                UIDA_Header dataGridHeader = new UIDA_Header(header);
                return dataGridHeader;
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
        new public UIDA_Group[] Groups
        {
            get
            {
                List<AutomationElement> allGroups = this.FindAll(ControlType.Group,
                    null, false, false, true);

                List<UIDA_Group> returnGroups = new List<UIDA_Group>();
                foreach (AutomationElement group in allGroups)
                {
                    UIDA_Group returnGroup = new UIDA_Group(group);
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
                    AutomationElement[] selectedItems = selectionPattern.Current.GetSelection();

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
            this.uiElement.TryGetCurrentPattern(SelectionPattern.Pattern, out selectionPatternObj);
            SelectionPattern selectionPattern = selectionPatternObj as SelectionPattern;

            return selectionPattern;
        }

        private GridPattern GetGridPattern()
        {
            object gridPatternObj = null;

            if (this.uiElement.TryGetCurrentPattern(GridPattern.Pattern, out gridPatternObj) == true)
            {
                GridPattern gridPattern = gridPatternObj as GridPattern;
                return gridPattern;
            }

            return null;
        }
    }
}
