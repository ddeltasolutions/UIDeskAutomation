using System;
using System.Collections.Generic;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
	/// <summary>
    /// This class represents a data item.
    /// </summary>
    public class UIDA_DataItem : ElementBase
    {
        /// <summary>
        /// Creates a UIDA_DataItem using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_DataItem(AutomationElement el)
        {
			this.uiElement = el;
		}

        /// <summary>
        /// Gets the selection state of a DataItem.
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
                
				UIDA_Header header = grid.Header;
				if (header == null)
				{
					throw new Exception("No header found");
				}
				
                UIDA_HeaderItem[] headerItems = header.Items;
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
}