using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// This class represents a tab control.
    /// </summary>
    public class TabCtrl : ElementBase
    {
        public TabCtrl(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets a collection with all Tab Items. It's like calling TabItems(null).
        /// </summary>
        public TabItem[] Items
        {
            get
            {
                List<AutomationElement> tabs = this.FindAll(ControlType.TabItem, null, false, false, true);
                List<TabItem> tabItems = new List<TabItem>();

                foreach (AutomationElement tab in tabs)
                {
                    TabItem tabItem = new TabItem(tab, this);
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
        /// <returns>TabItem element</returns>
        public TabItem TabItem(string name = null, bool searchDescendants = false,
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

            TabItem tabItem = new TabItem(returnElement, this);
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
        /// <returns>TabItem</returns>
        public TabItem TabItemAt(string name, int index, bool searchDescendants = false,
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

            TabItem tabItem = new TabItem(returnElement, this);
            return tabItem;
        }

        /// <summary>
        /// Returns a collection of TabItems that matches the search text (name), wildcards can be used.
        /// </summary>
        /// <param name="name">text of TabItem elements, use null to return all TabItems</param>
        /// <param name="searchDescendants">true is search deep through descendants, false is search through children, default false</param>
        /// <param name="caseSensitive">true if name search is done case sensitive, default true</param>
        /// <returns>TabItem elements</returns>
        public TabItem[] TabItems(string name = null, bool searchDescendants = false,
            bool caseSensitive = true)
        {
            List<AutomationElement> allTabItems = FindAll(ControlType.TabItem,
                name, searchDescendants, false, caseSensitive);

            List<TabItem> tabitems = new List<TabItem>();
            if (allTabItems != null)
            {
                foreach (AutomationElement crtEl in allTabItems)
                {
                    tabitems.Add(new TabItem(crtEl, this));
                }
            }
            return tabitems.ToArray();
        }

        /// <summary>
        /// Gets the currently selected Tab Item in this tab control.
        /// </summary>
        /// <returns>TabItem element</returns>
        public TabItem GetSelectedTabItem()
        {
            foreach (TabItem tabItem in this.Items)
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
            TabItem tabItem = TabItemAt(null, index, true);
            if (tabItem == null)
            {
                Engine.TraceInLogFile("TabItem not found");
                throw new Exception("TabItem not found");
                return;
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
            TabItem tabItem = TabItem(itemText, caseSensitive);
            if (tabItem == null)
            {
                Engine.TraceInLogFile("TabItem not found");
                throw new Exception("TabItem not found");
                return;
            }
            tabItem.Select();
        }
    }
}
