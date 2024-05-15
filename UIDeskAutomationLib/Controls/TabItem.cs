using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Tab Item control.
    /// </summary>
    public class UIDA_TabItem: ElementBase
    {
        public UIDA_TabItem(AutomationElement el, UIDA_TabCtrl parent)
        {
            this.uiElement = el;
            this.parent = parent;
        }
		
		public UIDA_TabItem(AutomationElement el)
		{
			this.uiElement = el;
			
			TreeWalker tw = TreeWalker.ControlViewWalker;
			AutomationElement uiparent = null;
			try
			{
				uiparent = tw.GetParent(this.uiElement);
			}
			catch { }
			
			while (uiparent != null)
			{
				if (uiparent.Current.ControlType == ControlType.Tab)
				{
					break;
				}
				
				uiparent = null;
				try
				{
					uiparent = tw.GetParent(uiparent);
				}
				catch { }
			}
			
			if (uiparent != null)
			{
				this.parent = new UIDA_TabCtrl(uiparent);
			}
		}

        private UIDA_TabCtrl parent = null;

        /// <summary>
        /// Returns true if the current tab item is selected, false otherwise.
        /// </summary>
        public bool IsSelected
        {
            get
            {
                object selectionItemPatternObj = null;

                if (this.uiElement.TryGetCurrentPattern(SelectionItemPattern.Pattern,
                    out selectionItemPatternObj) == true)
                {
                    SelectionItemPattern selectionItemPattern = selectionItemPatternObj as SelectionItemPattern;
                    if (selectionItemPattern != null)
                    {
                        try
                        {
                            return selectionItemPattern.Current.IsSelected;
                        }
                        catch (Exception ex)
                        {
                            Engine.TraceInLogFile("TreeItem.IsSelected failed: " + ex.Message);
                            throw new Exception("TreeItem.IsSelected failed: " + ex.Message);
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

            if (this.uiElement.TryGetCurrentPattern(SelectionItemPattern.Pattern, out selectionItemPatternObj) == true)
            {
                SelectionItemPattern selectionItemPattern = selectionItemPatternObj as SelectionItemPattern;
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
        /// Gets the zero based index of the current tab item.
        /// </summary>
        public int Index
        {
            get
            {
				if (this.parent == null)
				{
					return -1; // this tab item is not part of a tab control
				}
                UIDA_TabItem[] tabItems = this.parent.Items;

                for (int i = 0; i < tabItems.Length; i++)
                {
                    UIDA_TabItem tabItem = tabItems[i];

                    if (Helper.CompareAutomationElements(tabItem.uiElement, this.uiElement) == true)
                    {
                        return i;
                    }
                }

                return -1;
            }
        }
		
		/// <summary>
        /// Gets the text of the tab item. The same as calling GetText().
        /// </summary>
		public string Text
		{
			get
			{
				return this.GetText();
			}
		}
    }
}
