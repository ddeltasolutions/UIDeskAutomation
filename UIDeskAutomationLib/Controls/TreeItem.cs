using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a TreeView Item.
    /// </summary>
    public class UIDA_TreeItem: ElementBase
    {
        public UIDA_TreeItem(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets a collection of subitems of this TreeItem.
        /// </summary>
        /// <returns>UIDA_TreeItem array</returns>
        public UIDA_TreeItem[] SubItems
        {
            get
            {
                List<UIDA_TreeItem> tvItems = new List<UIDA_TreeItem>();

                List<AutomationElement> subItems =
                    this.FindAll(ControlType.TreeItem, null, false, false, true);

                foreach (AutomationElement el in subItems)
                {
                    UIDA_TreeItem tvItem = new UIDA_TreeItem(el);
                    tvItems.Add(tvItem);
                }

                return tvItems.ToArray();
            }
        }

        /// <summary>
        /// Expands a TreeItem.
        /// </summary>
        public void Expand()
        {
            object expandCollapsePatternObj = null;

            if (this.uiElement.TryGetCurrentPattern(ExpandCollapsePattern.Pattern,
                out expandCollapsePatternObj) == true)
            {
                ExpandCollapsePattern expandCollapsePattern =
                    expandCollapsePatternObj as ExpandCollapsePattern;

                if (expandCollapsePattern == null)
                {
                    Engine.TraceInLogFile("UIDA_TreeItem.Expand method failed");
                    throw new Exception("UIDA_TreeItem.Expand method failed");
                }

                expandCollapsePattern.Expand();
                return;
            }

            Engine.TraceInLogFile("UIDA_TreeItem.Expand method - cannot expand");
            throw new Exception("UIDA_TreeItem.Expand method - cannot expand");
        }

        /// <summary>
        /// Collapses a TreeItem.
        /// </summary>
        public void Collapse()
        {
            object expandCollapsePatternObj = null;

            if (this.uiElement.TryGetCurrentPattern(ExpandCollapsePattern.Pattern,
                out expandCollapsePatternObj) == true)
            {
                ExpandCollapsePattern expandCollapsePattern =
                    expandCollapsePatternObj as ExpandCollapsePattern;

                if (expandCollapsePattern == null)
                {
                    Engine.TraceInLogFile("UIDA_TreeItem.Collapse method failed");
                    throw new Exception("UIDA_TreeItem.Collapse method failed");
                }

                expandCollapsePattern.Collapse();
                return;
            }

            Engine.TraceInLogFile("UIDA_TreeItem.Collapse method - cannot collapse");
            throw new Exception("UIDA_TreeItem.Collapse method - cannot collapse");
        }
		
		/// <summary>
        /// Gets the ExpandCollapseState of this TreeItem.
        /// </summary>
        /// <returns>ExpandCollapseState</returns>
		public ExpandCollapseState ExpandCollapseState
		{
			get
			{
				object expandCollapsePatternObj = null;

				if (this.uiElement.TryGetCurrentPattern(ExpandCollapsePattern.Pattern,
					out expandCollapsePatternObj) == true)
				{
					ExpandCollapsePattern expandCollapsePattern =
						expandCollapsePatternObj as ExpandCollapsePattern;

					if (expandCollapsePattern == null)
					{
						Engine.TraceInLogFile("UIDA_TreeItem.ExpandCollapseState property failed");
						throw new Exception("UIDA_TreeItem.ExpandCollapseState property failed");
					}
					
					return expandCollapsePattern.Current.ExpandCollapseState;
				}
				
				Engine.TraceInLogFile("UIDA_TreeItem.ExpandCollapseState property not supported");
				throw new Exception("UIDA_TreeItem.ExpandCollapseState property not supported");
			}
		}

        /// <summary>
        /// Cycles through the toggle states (checked, unchecked, indeterminate).
        /// </summary>
        public void Toggle()
        {
			object togglePatternObj = null;

			if (this.uiElement.TryGetCurrentPattern(TogglePattern.Pattern,
				out togglePatternObj) == true)
			{
				TogglePattern togglePattern = togglePatternObj as TogglePattern;
				if (togglePattern == null)
				{
					Engine.TraceInLogFile("UIDA_TreeItem.Toggle() failed");
					throw new Exception("UIDA_TreeItem.Toggle() failed");
				}
				
				try
				{
					togglePattern.Toggle();
					return;
				}
				catch (Exception ex)
				{
					Engine.TraceInLogFile("UIDA_TreeItem.Toggle() error: " + ex.Message);
					throw new Exception("UIDA_TreeItem.Toggle() error: " + ex.Message);
				} 
			}
			
			Engine.TraceInLogFile("UIDA_TreeItem.Toggle() failed: TogglePattern not supported");
			throw new Exception("UIDA_TreeItem.Toggle() failed: TogglePattern not supported");
        }

        /// <summary>
        /// Selects the current tree item and deselects all others selected tree items.
        /// </summary>
        public void Select()
        {
            object selectionItemPatternObj = null;

            if (this.uiElement.TryGetCurrentPattern(SelectionItemPattern.Pattern,
                out selectionItemPatternObj) == true)
            {
                SelectionItemPattern selectionItemPattern =
                    selectionItemPatternObj as SelectionItemPattern;

                if (selectionItemPattern == null)
                {
                    Engine.TraceInLogFile("UIDA_TreeItem.Select() method failed");
                    throw new Exception("UIDA_TreeItem.Select() method failed");
                }

                selectionItemPattern.Select();
                return;
            }

            Engine.TraceInLogFile("UIDA_TreeItem.Select() method failed");
            throw new Exception("UIDA_TreeItem.Select() method failed");
        }

        /// <summary>
        /// Brings the current TreeView item into viewable area of the parent Tree control.
        /// </summary>
        public void BringIntoView()
        {
            object scrollItemPatternObj = null;

            if (this.uiElement.TryGetCurrentPattern(ScrollItemPattern.Pattern,
                out scrollItemPatternObj) == true)
            {
                ScrollItemPattern scrollItemPattern =
                    scrollItemPatternObj as ScrollItemPattern;

                if (scrollItemPattern == null)
                {
                    Engine.TraceInLogFile("UIDA_TreeItem.BringIntoView method failed");
                    throw new Exception("UIDA_TreeItem.BringIntoView method failed");
                }

                try
                {
                    scrollItemPattern.ScrollIntoView();
                    return;
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("UIDA_TreeItem.BringIntoView method failed: " + 
                        ex.Message);
                    throw new Exception("UIDA_TreeItem.BringIntoView method failed: " +
                        ex.Message);
                }
            }

            Engine.TraceInLogFile("UIDA_TreeItem.BringIntoView method failed");
            throw new Exception("UIDA_TreeItem.BringIntoView method failed");
        }
		
		/// <summary>
        /// Gets or sets the checked state of the tree item if supported.
        /// </summary>
		/// <returns>true if tree item is checked, false otherwise</returns>
        public bool IsChecked
        {
            get
            {
                object togglePatternObj = null;

                if (this.uiElement.TryGetCurrentPattern(TogglePattern.Pattern,
                    out togglePatternObj) == true)
                {
                    TogglePattern togglePattern = togglePatternObj as TogglePattern;

                    if (togglePattern == null)
                    {
                        Engine.TraceInLogFile("UIDA_TreeItem.Checked failed");
                        throw new Exception("UIDA_TreeItem.Checked failed");
                    }

                    try
                    {
                        if (togglePattern.Current.ToggleState == ToggleState.On)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    catch (Exception ex)
                    { 
                        Engine.TraceInLogFile("UIDA_TreeItem.Checked failed: " + ex.Message);
                        throw new Exception("UIDA_TreeItem.Checked failed: " + ex.Message);
                    }
                }

                Engine.TraceInLogFile("UIDA_TreeItem.Checked not supported");
                throw new Exception("UIDA_TreeItem.Checked not supported");
            }

            set
            {
                object togglePatternObj = null;

                if (this.uiElement.TryGetCurrentPattern(TogglePattern.Pattern,
                    out togglePatternObj) == true)
                {
                    TogglePattern togglePattern = togglePatternObj as TogglePattern;

                    if (togglePattern == null)
                    {
                        if (value == true)
                        {
                            Engine.TraceInLogFile("Cannot check tree item");
                            throw new Exception("Cannot check tree item");
                        }
                        else
                        {
                            Engine.TraceInLogFile("Cannot uncheck tree item");
                            throw new Exception("Cannot uncheck tree item");
                        }
                    }

                    if (value == true)
                    {
                        // try to check list item
                        try
                        {
                            if (togglePattern.Current.ToggleState != ToggleState.On)
                            {
                                togglePattern.Toggle();
                            }

                            if (togglePattern.Current.ToggleState != ToggleState.On)
                            {
                                //this.SimulateDoubleClick();
                                togglePattern.Toggle();
                            }

                            return;
                        }
                        catch (Exception ex)
                        {
                            Engine.TraceInLogFile("Cannot check tree item: " + ex.Message);
                            throw new Exception("Cannot check tree item: " + ex.Message);
                        }
                    }
                    else
                    {
                        // try to uncheck
                        try
                        {
                            if (togglePattern.Current.ToggleState != ToggleState.Off)
                            {
                                togglePattern.Toggle();
                            }

                            if (togglePattern.Current.ToggleState != ToggleState.Off)
                            {
                                //this.SimulateDoubleClick();
								togglePattern.Toggle();
                            }

                            return;
                        }
                        catch (Exception ex)
                        {
                            Engine.TraceInLogFile("Cannot uncheck tree item" + ex.Message);
                            throw new Exception("Cannot uncheck tree item" + ex.Message);
                        }
                    }
                }

                if (value == true)
                {
                    Engine.TraceInLogFile("Cannot check tree item");
                    throw new Exception("Cannot check tree item");
                }
                else
                { 
                    Engine.TraceInLogFile("Cannot uncheck tree item");
                    throw new Exception("Cannot uncheck tree item");
                }
            }
        }
    }
}
