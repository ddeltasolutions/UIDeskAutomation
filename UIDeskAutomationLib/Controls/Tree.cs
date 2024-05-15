using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a Tree control.
    /// </summary>
    public class UIDA_Tree: ElementBase
    {
        public UIDA_Tree(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets the tree root item.
        /// </summary>
        /// <returns>UIDA_TreeItem element</returns>
        public UIDA_TreeItem GetRoot()
        {
            AutomationElement rootElement = this.FindFirst(ControlType.TreeItem, null,
                false, false, true);

            if (rootElement == null)
            {
                Engine.TraceInLogFile("GetRoot() method - root element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GetRoot() method - root element not found");
                }
                else
                {
                    return null;
                }
            }

            UIDA_TreeItem root = new UIDA_TreeItem(rootElement);
            return root;
        }
		
		private AutomationEventHandler UIAeventHandler = null;
		
		/// <summary>
        /// Delegate for Element Selected event
        /// </summary>
		/// <param name="sender">The tree control that sent the event.</param>
		/// <param name="selectedItem">the new selected tree item.</param>
		public delegate void ElementSelected(UIDA_Tree sender, UIDA_TreeItem selectedItem);
		internal ElementSelected ElementSelectedHandler = null;
		
		/// <summary>
        /// Attaches/detaches a handler to element selected event
        /// </summary>
		public event ElementSelected ElementSelectedEvent
		{
			add
			{
				try
				{
					if (this.ElementSelectedHandler == null)
					{
						UIAeventHandler = new AutomationEventHandler(OnUIAutomationEvent);
			
						Automation.AddAutomationEventHandler(SelectionItemPattern.ElementSelectedEvent,
							base.uiElement, TreeScope.Subtree, UIAeventHandler);
					}
					
					this.ElementSelectedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.ElementSelectedHandler -= value;
				
					if (this.ElementSelectedHandler == null)
					{
						if (this.UIAeventHandler == null)
						{
							return;
						}
						
						System.Threading.Tasks.Task.Run(() => 
						{
							try
							{
								Automation.RemoveAutomationEventHandler(SelectionItemPattern.ElementSelectedEvent, 
									base.uiElement, this.UIAeventHandler);
								UIAeventHandler = null;
							}
							catch { }
						}).Wait(5000);
					}
				}
				catch {}
			}
		}
		
		private void OnUIAutomationEvent(object sender, AutomationEventArgs e)
		{
			if (e.EventId == SelectionItemPattern.ElementSelectedEvent && this.ElementSelectedHandler != null)
			{
				AutomationElement sourceElement = sender as AutomationElement;
				if (sourceElement != null)
				{
					ElementSelectedHandler(this, new UIDA_TreeItem(sourceElement));
				}
				else
				{
					ElementSelectedHandler(this, null);
				}
			}
		}
    }
}
