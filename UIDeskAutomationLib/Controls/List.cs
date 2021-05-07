using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a List UI element
    /// </summary>
    public class List: ElementBase
    {
		/// <summary>
        /// Creates a List using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public List(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets a collection of all list items. It is like calling ListItems(null).
        /// </summary>
        public ListItem[] Items
        {
            get
            {
                if (uiElement.Current.FrameworkId != "WPF")
                {
                    List<AutomationElement> allListItems = 
                        this.FindAll(ControlType.ListItem, null, false, false, true);

                    List<ListItem> returnCollection = new List<ListItem>();

                    foreach (AutomationElement listItemEl in allListItems)
                    {
                        ListItem listItem = new ListItem(listItemEl);
                        returnCollection.Add(listItem);
                    }

                    return returnCollection.ToArray();
                }
                else
                {
                    List<ListItem> returnRows = new List<ListItem>();
                    
                    object objectPattern = null;
                    this.uiElement.TryGetCurrentPattern(ItemContainerPattern.Pattern, out objectPattern);
                    ItemContainerPattern itemContainerPattern = objectPattern as ItemContainerPattern;
                    if (itemContainerPattern == null)
                    {
                        List<AutomationElement> allListItems = 
                            this.FindAll(ControlType.ListItem, null, false, false, true);
                        foreach (AutomationElement listItemEl in allListItems)
                        {
                            ListItem listItem = new ListItem(listItemEl);
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
                            ListItem listItem = new ListItem(crt);
                            returnRows.Add(listItem);
                        }
                    }
                    while (crt != null);
                    
                    return returnRows.ToArray();
                }
            }
        }

        /// <summary>
        /// Selects all list items in the current list
        /// </summary>
        public void SelectAll()
        { 
            ListItem[] allListItems = this.Items;

            foreach (ListItem listItem in allListItems)
            {
                if (listItem.IsSelected == false)
                {
                    listItem.AddToSelection();
                }
            }
        }

        /// <summary>
        /// Returns a collection with all selected list items in the current list
        /// </summary>
        public ListItem[] SelectedItems
        { 
            get
            {
                System.Collections.Generic.List<ListItem> selectedItems =
                    new List<ListItem>();

                ListItem[] allListItems = this.Items;

                foreach (ListItem listItem in allListItems)
                {
                    if (listItem.IsSelected == true)
                    {
                        selectedItems.Add(listItem);
                    }
                }

                return selectedItems.ToArray();
            }
        }

        /// <summary>
        /// Clears all selections in the current list.
        /// </summary>
        public void ClearAllSelection()
        {
            ListItem[] selectedItems = this.SelectedItems;

            foreach (ListItem selectedListItem in selectedItems)
            {
                selectedListItem.RemoveFromSelection();
            }
        }
        
        private AutomationElement[] GetItemsByText(string text, bool all)
        {
            object objectPattern = null;
            this.uiElement.TryGetCurrentPattern(ItemContainerPattern.Pattern, out objectPattern);
            ItemContainerPattern itemContainerPattern = objectPattern as ItemContainerPattern;
            
            if (itemContainerPattern != null)
            {
                List<AutomationElement> items = new List<AutomationElement>();
                if (all == false)
                {
                    AutomationElement item = itemContainerPattern.FindItemByProperty(null, AutomationElement.NameProperty, text);
                    if (item != null)
                    {
                        items.Add(item);
                    }
                }
                else
                {
                    AutomationElement crt = null;
                    do
                    {
                        crt = itemContainerPattern.FindItemByProperty(crt, AutomationElement.NameProperty, text);
                        if (crt != null)
                        {
                            items.Add(crt);
                        }
                    }
                    while (crt != null);
                }
                return items.ToArray();
            }
            return null;
        }
        
        private AutomationElement GetWPFListItem(int index)
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
                if (crt != null)
                {
                    //Engine.TraceInLogFile("Element name: " + crt.Current.Name);
                }
                else
                {
                    Engine.TraceInLogFile("list item index too big");
                    break;
                }
                index--;
            }
            while (index != 0);
            
			return (index == 0 ? crt : null);
        }
        
        /// <summary>
        /// Selects an item in a List by the item index. Other selected items will be deselected.
        /// </summary>
        /// <param name="index">item index, starts with 1</param>
        public void Select(int index)
        {
            if (uiElement.Current.FrameworkId == "WPF")
            {
                AutomationElement item = GetWPFListItem(index);
                if (item != null)
                {
                    ListItem listItemWPF = new ListItem(item);
                    listItemWPF.Select();
                    return;
                }
            }
            
            ListItem listItem = ListItemAt(null, index, true);
            if (listItem == null)
            {
                Engine.TraceInLogFile("Item not found");
                throw new Exception("Item not found");
            }
            listItem.Select();
        }
        
        /// <summary>
        /// Selects an item (or more items) in a List by the item text. Other selected items will be deselected.
        /// </summary>
        /// <param name="itemText">Item text. Wildcards can be used.</param>
        /// <param name="selectAll">true to select all items matching the given text, false to select only the first item matching the given text</param>
        /// <param name="caseSensitive">true if the item text search is done case sensitive</param>
        public void Select(string itemText = null, bool selectAll = true, bool caseSensitive = true)
        {
            if (uiElement.Current.FrameworkId == "WPF")
            {
                AutomationElement[] items = GetItemsByText(itemText, selectAll);
                if (items != null && items.Length > 0)
                {
                    ListItem listItem = new ListItem(items[0]);
                    listItem.Select();
                    for (int i = 1; i < items.Length; i++)
                    {
                        listItem = new ListItem(items[i]);
                        listItem.AddToSelection();
                    }
                    return;
                }
                else
                {
                    Engine.TraceInLogFile("FindItemByProperty didn't find the item");
                }
            }
            
            if (selectAll == false)
            {
                ListItem listItem = ListItem(itemText, true, caseSensitive);
                if (listItem == null)
                {
                    Engine.TraceInLogFile("Item not found");
                    throw new Exception("Item not found");
                }
                
                listItem.Select();
            }
            else
            {
                ListItem[] items = ListItems(itemText, true, caseSensitive);
                if (items.Length == 0)
                {
                    return;
                }
                items[0].Select();
                for (int i = 1; i < items.Length; i++)
                {
                    items[i].AddToSelection();
                }
            }
        }
        
        /// <summary>
        /// Adds an item to selection in a List by the item index.
        /// </summary>
        /// <param name="index">item index, starts with 1</param>
        public void AddToSelection(int index)
        {
            if (uiElement.Current.FrameworkId == "WPF")
            {
                AutomationElement item = GetWPFListItem(index);
                if (item != null)
                {
                    ListItem listItemWPF = new ListItem(item);
                    listItemWPF.AddToSelection();
                    return;
                }
            }
            
            ListItem listItem = ListItemAt(null, index, true);
            if (listItem == null)
            {
                Engine.TraceInLogFile("Item not found");
                throw new Exception("Item not found");
            }
            listItem.AddToSelection();
        }
        
        /// <summary>
        /// Adds an item (or more items) to selection in a List by the item text.
        /// </summary>
        /// <param name="itemText">Item text. Wildcards can be used.</param>
        /// <param name="selectAll">true to add to selection all items matching the given text, false to add to selection only the first item matching the given text</param>
        /// <param name="caseSensitive">true if the item text search is done case sensitive</param>
        public void AddToSelection(string itemText = null, bool selectAll = true, bool caseSensitive = true)
        {
            if (uiElement.Current.FrameworkId == "WPF")
            {
                AutomationElement[] items = GetItemsByText(itemText, selectAll);
                if (items != null && items.Length > 0)
                {
                    ListItem listItem = new ListItem(items[0]);
                    listItem.AddToSelection();
                    for (int i = 1; i < items.Length; i++)
                    {
                        listItem = new ListItem(items[i]);
                        listItem.AddToSelection();
                    }
                    return;
                }
                else
                {
                    Engine.TraceInLogFile("FindItemByProperty didn't find the item");
                }
            }
            
            if (selectAll == false)
            {
                ListItem listItem = ListItem(itemText, true, caseSensitive);
                if (listItem == null)
                {
                    Engine.TraceInLogFile("Item not found");
                    throw new Exception("Item not found");
                }
                
                listItem.AddToSelection();
            }
            else
            {
                ListItem[] items = ListItems(itemText, true, caseSensitive);

                foreach (ListItem item in items)
                {
                    item.AddToSelection();
                }
            }
        }
        
        /// <summary>
        /// Removes an item from selection in a List by the item index.
        /// </summary>
        /// <param name="index">item index, starts with 1</param>
        public void RemoveFromSelection(int index)
        {
            if (uiElement.Current.FrameworkId == "WPF")
            {
                AutomationElement item = GetWPFListItem(index);
                if (item != null)
                {
                    ListItem listItemWPF = new ListItem(item);
                    listItemWPF.RemoveFromSelection();
                    return;
                }
            }
            
            ListItem listItem = ListItemAt(null, index, true);
            if (listItem == null)
            {
                Engine.TraceInLogFile("Item not found");
                throw new Exception("Item not found");
            }
            listItem.RemoveFromSelection();
        }
        
        /// <summary>
        /// Removes an item (or more items) from selection in a List by the item text.
        /// </summary>
        /// <param name="itemText">Item text. Wildcards can be used.</param>
        /// <param name="all">true to remove from selection all items matching the given text, false to remove from selection only the first item matching the given text</param>
        /// <param name="caseSensitive">true if the item text search is done case sensitive</param>
        public void RemoveFromSelection(string itemText = null, bool all = true, bool caseSensitive = true)
        {
            if (uiElement.Current.FrameworkId == "WPF")
            {
                AutomationElement[] items = GetItemsByText(itemText, all);
                if (items != null && items.Length > 0)
                {
                    ListItem listItem = new ListItem(items[0]);
                    listItem.RemoveFromSelection();
                    for (int i = 1; i < items.Length; i++)
                    {
                        listItem = new ListItem(items[i]);
                        listItem.RemoveFromSelection();
                    }
                    return;
                }
                else
                {
                    Engine.TraceInLogFile("FindItemByProperty didn't find the item");
                }
            }
            
            if (all == false)
            {
                ListItem listItem = ListItem(itemText, true, caseSensitive);
                if (listItem == null)
                {
                    Engine.TraceInLogFile("Item not found");
                    throw new Exception("Item not found");
                }
                
                listItem.RemoveFromSelection();
            }
            else
            {
                ListItem[] items = ListItems(itemText, true, caseSensitive);

                foreach (ListItem item in items)
                {
                    item.RemoveFromSelection();
                }
            }
        }
    }
}
