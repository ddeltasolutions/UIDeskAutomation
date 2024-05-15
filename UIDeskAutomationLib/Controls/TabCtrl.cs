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
