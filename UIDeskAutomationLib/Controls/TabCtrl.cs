using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// This class represents a tab control.
    /// </summary>
    public class UIDA_TabCtrl : ElementBase
    {
		/// <summary>
        /// Creates a UIDA_TabCtrl using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_TabCtrl(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets a collection with all Tab Items. It's like calling TabItems(null).
        /// </summary>
        public UIDA_TabItem[] Items
        {
            get
            {
                List<AutomationElement> tabs = this.FindAll(ControlType.TabItem, null, false, false, true);
                List<UIDA_TabItem> tabItems = new List<UIDA_TabItem>();

                foreach (AutomationElement tab in tabs)
                {
                    UIDA_TabItem tabItem = new UIDA_TabItem(tab, this);
                    tabItems.Add(tabItem);
                }

                return tabItems.ToArray();
            }
        }

        /// <summary>
        /// Searches a TabItem page in the current tab control
        /// </summary>
        /// <param name="name">name of tab item</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is case sensitive, default true</param>
        /// <returns>UIDA_TabItem element</returns>
        public UIDA_TabItem TabItem(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            AutomationElement returnElement = this.FindFirst(ControlType.TabItem,
                name, searchDescendants, false, caseSensitive);

            if (returnElement == null)
            {
                Engine.TraceInLogFile("TabItem method - TabItem element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TabItem method - TabItem element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_TabItem tabItem = new UIDA_TabItem(returnElement, this);
            return tabItem;
        }

        /// <summary>
        /// Searches for a tab item among the children of the current tab control
        /// or its descendants.
        /// </summary>
        /// <param name="name">name of tab item</param>
        /// <param name="index">tab item index</param>
        /// <param name="searchDescendants">search descendants, default false</param>
        /// <param name="caseSensitive">search name with case sensitive criteria</param>
        /// <returns>UIDA_TabItem</returns>
        public UIDA_TabItem TabItemAt(string name, int index, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("TabItemAt method - index cannot be less than zero");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TabItemAt method - index cannot be less than zero");
                }
                else
                {
                    return null;
                }
            }

            AutomationElement returnElement = null;

            Errors error = this.FindAt(ControlType.TabItem, name, index, searchDescendants,
                false, caseSensitive, out returnElement);

            if (error == Errors.ElementNotFound)
            {
                Engine.TraceInLogFile("TabItemAt method - button element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TabItemAt method - button element not found");
                }
                else
                {
                    return null;
                }
            }
            else if (error == Errors.IndexTooBig)
            {
                Engine.TraceInLogFile("TabItemAt method - index too big");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("TabItemAt method - index too big");
                }
                else
                {
                    return null;
                }
            }

            UIDA_TabItem tabItem = new UIDA_TabItem(returnElement, this);
            return tabItem;
        }

        /// <summary>
        /// Returns a collection of TabItems that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of TabItem elements, use null to return all TabItems</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>UIDA_TabItem elements</returns>
        public UIDA_TabItem[] TabItems(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<AutomationElement> allTabItems = FindAll(ControlType.TabItem,
                name, searchDescendants, false, caseSensitive);

            List<UIDA_TabItem> tabitems = new List<UIDA_TabItem>();
            if (allTabItems != null)
            {
                foreach (AutomationElement crtEl in allTabItems)
                {
                    tabitems.Add(new UIDA_TabItem(crtEl, this));
                }
            }
            return tabitems.ToArray();
        }

        /// <summary>
        /// Gets the currently selected Tab Item in this tab control.
        /// </summary>
        /// <returns>UIDA_TabItem element</returns>
        public UIDA_TabItem GetSelectedTabItem()
        {
            foreach (UIDA_TabItem tabItem in this.Items)
            {
                if (tabItem.IsSelected)
                {
                    return tabItem;
                }
            }

            Engine.TraceInLogFile("TreeItem.GetSelectedItem() method: No tab item is selected");
            return null;
        }
        
        /// <summary>
        /// Selects a TabItem in a TabCtrl by the tab item index.
        /// </summary>
        /// <param name="index">tab item index, starts with 1</param>
        public void Select(int index)
        {
            UIDA_TabItem tabItem = TabItemAt(null, index, true);
            if (tabItem == null)
            {
                Engine.TraceInLogFile("TabItem not found");
                throw new Exception("TabItem not found");
            }
            tabItem.Select();
        }
        
        /// <summary>
        /// Selects a TabItem in a TabCtrl by the tab item text. Wildcards can be used.
        /// </summary>
        /// <param name="itemText">tab item text</param>
        /// <param name="caseSensitive">true if the tab item text search is done case sensitive</param>
        public void Select(string itemText = null, bool caseSensitive = true)
        {
            UIDA_TabItem tabItem = TabItem(itemText, caseSensitive);
            if (tabItem == null)
            {
                Engine.TraceInLogFile("TabItem not found");
                throw new Exception("TabItem not found");
            }
            tabItem.Select();
        }
    }
}
