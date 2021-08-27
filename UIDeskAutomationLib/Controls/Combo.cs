using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a ComboBox UI element.
    /// </summary>
    public class UIDA_ComboBox: ElementBase
    {
        /// <summary>
        /// Creates a UIDA_ComboBox using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_ComboBox(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets a collection with all ComboBox items.
        /// </summary>
        public UIDA_ListItem[] Items
        {
            get
            {
                if (uiElement.Current.FrameworkId != "WPF")
                {
                    List<AutomationElement> allItems =
                        this.FindAll(ControlType.ListItem, null, true, false, true);

                    List<UIDA_ListItem> returnCollection = new List<UIDA_ListItem>();

                    foreach (AutomationElement el in allItems)
                    {
                        UIDA_ListItem listItem = new UIDA_ListItem(el);
                        returnCollection.Add(listItem);
                    }

                    return returnCollection.ToArray();
                }
                else
                {
                    List<UIDA_ListItem> returnRows = new List<UIDA_ListItem>();
                    
                    object objectPattern = null;
                    this.uiElement.TryGetCurrentPattern(ItemContainerPattern.Pattern, out objectPattern);
                    ItemContainerPattern itemContainerPattern = objectPattern as ItemContainerPattern;
                    if (itemContainerPattern == null)
                    {
                        List<AutomationElement> allListItems = 
                            this.FindAll(ControlType.ListItem, null, true, false, true);
                        foreach (AutomationElement listItemEl in allListItems)
                        {
                            UIDA_ListItem listItem = new UIDA_ListItem(listItemEl);
                            returnRows.Add(listItem);
                        }
                        return returnRows.ToArray();
                    }
                    
                    AutomationElement crt = null;
                    do
                    {
                        crt = itemContainerPattern.FindItemByProperty(crt, null, null);
                        if (crt != null)
                        {
                            UIDA_ListItem listItem = new UIDA_ListItem(crt);
                            returnRows.Add(listItem);
                        }
                    }
                    while (crt != null);
                    
                    return returnRows.ToArray();
                }
            }
        }

        /// <summary>
        /// Searches for a combobox item.
        /// </summary>
        /// <param name="name">name of combobox item</param>
        /// <param name="caseSensitive">true is name search is done case sensitive, default true</param>
        /// <returns>ComboBox item</returns>
        public UIDA_ListItem ListItem(string name = null, bool caseSensitive = true)
        {
            AutomationElement returnElement = this.FindFirst(ControlType.ListItem, name,
                true, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("Combo::ListItem method - list item element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Combo::ListItem method - list item element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ListItem listItem = new UIDA_ListItem(returnElement);
            return listItem;
        }

        /// <summary>
        /// Searches for a ComboBox item with specified name at specified index.
        /// </summary>
        /// <param name="name">text of ComboBox item</param>
        /// <param name="index">index</param>
        /// <param name="caseSensitive">true is name search is done case sensitive, false otherwise, default true</param>
        /// <returns>UIDA_ListItem element of ComboBox</returns>
        public UIDA_ListItem ListItemAt(string name, int index, bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("Combo::ListItemAt method - index cannot be negative");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Combo::ListItemAt method - index cannot be negative");
                }
                else
                {
                    return null;
                }
            }

            AutomationElement returnElement = null;

            Errors error = this.FindAt(ControlType.ListItem, name, index, true,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("Combo::ListItemAt method - ComboBox item element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Combo::ListItemAt method - ComboBox item element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("Combo::ListItemAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("Combo::ListItemAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_ListItem listItem = new UIDA_ListItem(returnElement);
            return listItem;
        }

        /// <summary>
        /// Returns a collection of ListItems that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of ListItem elements, use null to return all ListItems</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_ListItem elements</returns>
        new public UIDA_ListItem[] ListItems(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<AutomationElement> allListItems = FindAll(ControlType.ListItem,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_ListItem> listitems = new List<UIDA_ListItem>();
            if (allListItems != null)
            {
                foreach (AutomationElement crtEl in allListItems)
                {
                    listitems.Add(new UIDA_ListItem(crtEl));
                }
            }
            return listitems.ToArray();
        }

        /// <summary>
        /// Sets a text for a drop-down ComboBox.
        /// </summary>
        /// <param name="text">text to set</param>
        public void SetText(string text)
        {
            UIDA_Edit edit = this.Edit();

            if (edit == null)
            {
                Engine.TraceInLogFile("ComboBox::SetText failed: cannot set text, ComboBox should have DropDown style");
                throw new Exception("ComboBox::SetText failed: cannot set text. ComboBox should have DropDown style");
            }

            edit.SetText(text);
        }

        /// <summary>
        /// Gets the selected ComboBox item.
        /// </summary>
        public UIDA_ListItem SelectedItem
        {
            get
            {
                object objectPattern = null;
                this.uiElement.TryGetCurrentPattern(SelectionPattern.Pattern, out objectPattern);
                SelectionPattern selectionPattern = objectPattern as SelectionPattern;
                
                if (selectionPattern == null)
                {
                    UIDA_ListItem[] allItems = this.Items;

                    foreach (UIDA_ListItem item in allItems)
                    {
                        if (item.IsSelected)
                        {
                            return item;
                        }
                    }
                }
                else
                {
                    // SelectionPattern is supported
                    AutomationElement[] selection = selectionPattern.Current.GetSelection();
                    if (selection.Length >= 1)
                    {
                        UIDA_ListItem listItem = new UIDA_ListItem(selection[0]);
                        return listItem;
                    }
                }
                
                // no item is selected
                return null;
            }
        }
        
        /// <summary>
        /// Expands the combobox.
        /// </summary>
        public void Expand()
        {
            object objectPattern = null;
            this.uiElement.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out objectPattern);
            ExpandCollapsePattern expandCollapsePattern = objectPattern as ExpandCollapsePattern;
            
            if (expandCollapsePattern != null)
            {
                expandCollapsePattern.Expand();
            }
        }
        
        /// <summary>
        /// Collapses the combobox.
        /// </summary>
        public void Collapse()
        {
            object objectPattern = null;
            this.uiElement.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out objectPattern);
            ExpandCollapsePattern expandCollapsePattern = objectPattern as ExpandCollapsePattern;
            
            if (expandCollapsePattern != null)
            {
                expandCollapsePattern.Collapse();
            }
        }
        
        /// <summary>
        /// Selects an item in a ComboBox by the item index.
        /// </summary>
        /// <param name="index">item index, starts with 1</param>
        public void Select(int index)
        {
            string fid = this.uiElement.Current.FrameworkId;
            if (fid == "WPF")
            {
                Expand();
            }
            
            UIDA_ListItem listItem = ListItemAt(null, index);
            if (listItem == null)
            {
                Engine.TraceInLogFile("Item not found");
                throw new Exception("Item not found");
            }
            listItem.Select();
            
            if (fid == "WPF")
            {
                Collapse();
            }
        }
        
        /// <summary>
        /// Selects an item in a ComboBox by the item text. Wildcards can be used.
        /// </summary>
        /// <param name="itemText">item text</param>
        /// <param name="caseSensitive">true if the item text search is done case sensitive</param>
        public void Select(string itemText = null, bool caseSensitive = true)
        {
            string fid = this.uiElement.Current.FrameworkId;
            if (fid == "WPF")
            {
                Expand();
            }
            
            UIDA_ListItem listItem = ListItem(itemText, caseSensitive);
            if (listItem == null)
            {
                Engine.TraceInLogFile("Item not found");
                throw new Exception("Item not found");
            }
            
            listItem.Select();
            
            if (fid == "WPF")
            {
                Collapse();
            }
        }
    }
}
