using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Represents a Tab Item control.
    /// </summary>
    public class TabItem: ElementBase
    {
        public TabItem(AutomationElement el, TabCtrl parent)
        {
            this.uiElement = el;
            this.parent = parent;
        }

        private TabCtrl parent = null;

        /// <summary>
        /// Returns true is current tab is selected, false otherwise.
        /// </summary>
        public bool IsSelected
        {
            get
            {
                object selectionItemPatternObj = null;

                if (this.uiElement.TryGetCurrentPattern(SelectionItemPattern.Pattern,
                    out selectionItemPatternObj) == true)
                {
                    SelectionItemPattern selectionItemPattern =
                        selectionItemPatternObj as SelectionItemPattern;

                    if (selectionItemPattern != null)
                    {
                        try
                        {
                            return selectionItemPattern.Current.IsSelected;
                        }
                        catch (Exception ex)
                        {
                            Engine.TraceInLogFile("TreeItem.IsSelected failed: " +
                                ex.Message);

                            throw new Exception("TreeItem.IsSelected failed: " +
                                ex.Message);
                        }
                    }
                }

                Engine.TraceInLogFile("TreeItem.IsSelected failed");
                throw new Exception("TreeItem.IsSelected failed");
            }
        }

        /// <summary>
        /// Selects the current tab item.
        /// </summary>
        public void Select()
        {
            object selectionItemPatternObj = null;

            if (this.uiElement.TryGetCurrentPattern(SelectionItemPattern.Pattern,
                out selectionItemPatternObj) == true)
            {
                SelectionItemPattern selectionItemPattern =
                    selectionItemPatternObj as SelectionItemPattern;

                if (selectionItemPattern != null)
                {
                    try
                    {
                        selectionItemPattern.Select();

                        return;
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile("TabItem.Select() failed: " + ex.Message);
                        throw new Exception("TabItem.Select() failed: " + ex.Message);
                    }
                }
            }

            Engine.TraceInLogFile("TabItem.Select() method failed");
            throw new Exception("TabItem.Select() method failed");
        }

        /// <summary>
        /// Zero based index of the current tab item.
        /// </summary>
        public int Index
        {
            get
            {
                TabItem[] tabItems = this.parent.Items;

                for (int i = 0; i < tabItems.Length; i++)
                {
                    TabItem tabItem = tabItems[i];

                    if (Helper.CompareAutomationElements(
                        tabItem.uiElement, this.uiElement) == true)
                    {
                        return i;
                    }
                }

                return -1;
            }
        }
    }
}
