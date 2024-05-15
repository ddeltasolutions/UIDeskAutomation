using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a RadioButton UI element.
    /// </summary>
    public class UIDA_RadioButton: ElementBase
    {
        public UIDA_RadioButton(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets the selected state of the radio button.
        /// </summary>
        /// <returns>true - is selected, false otherwise</returns>
        public bool IsSelected
        {
			get
			{
				object selectionItemPatternObj = null;
				if (this.uiElement.TryGetCurrentPattern(SelectionItemPattern.Pattern,
					out selectionItemPatternObj) == true)
				{
					SelectionItemPattern selectionItemPattern = selectionItemPatternObj as SelectionItemPattern;
					if (selectionItemPattern == null)
					{
						Engine.TraceInLogFile("RadioButton.GetIsSelected() - method failed");
						throw new Exception("RadioButton.GetIsSelected() - method failed");
					}

					bool isSelected = false;

					try
					{
						isSelected = selectionItemPattern.Current.IsSelected;
						return isSelected;
					}
					catch { }
				}

				Engine.TraceInLogFile("RadioButton.GetIsSelected() - method failed");
				throw new Exception("RadioButton.GetIsSelected() - method failed");
			}
        }

        /// <summary>
        /// Selects a radio button.
        /// </summary>
        public void Select()
        {
			this.Click();
			//Engine.GetInstance().Sleep(100);
			
			/*if (this.uiElement.Current.FrameworkId == "WPF")
			{
				// for WPF use mouse click because SelectionItemPattern.Select() is not calling the radio button's Click event handler
				this.Click();
				return;
			}
			
            object selectionItemPatternObj = null;

            if (this.uiElement.TryGetCurrentPattern(SelectionItemPattern.Pattern,
                out selectionItemPatternObj) == true)
            {
                SelectionItemPattern selectionItemPattern =
                    selectionItemPatternObj as SelectionItemPattern;

                if (selectionItemPattern == null)
                {
                    Engine.TraceInLogFile("RadioButton.Select() - method failed");
                    throw new Exception("RadioButton.Select() - method failed");
                }

                try
                {
                    selectionItemPattern.Select();
                    return;
                }
                catch { }
            }

            Engine.TraceInLogFile("RadioButton.Select() - method failed");
            throw new Exception("RadioButton.Select() - method failed"); */
        }
		
		/// <summary>
        /// Gets the text of the radio button
        /// </summary>
		public string Text
		{
			get
			{
				return this.GetText();
			}
		}
		
		private AutomationPropertyChangedEventHandler UIAPropChangedEventHandler = null;
		private AutomationEventHandler UIAeventHandler = null;
		
		/// <summary>
        /// Delegate for State Changed event
        /// </summary>
		/// <param name="sender">The radio button that sent the event.</param>
		/// <param name="isChecked">true if the radio button is checked, false if it's unchecked.</param>
		public delegate void StateChanged(UIDA_RadioButton sender, bool isChecked);
		internal StateChanged StateChangedHandler = null;
		
		/// <summary>
        /// Attaches/detaches a handler to state changed event
        /// </summary>
		public event StateChanged StateChangedEvent
		{
			add
			{
				try
				{
					if (this.StateChangedHandler == null)
					{
						string cfid = base.uiElement.Current.FrameworkId;
						if (cfid == "Win32")
						{
							UIAeventHandler = new AutomationEventHandler(OnUIAutomationEvent);
						
							Automation.AddAutomationEventHandler(SelectionItemPattern.ElementSelectedEvent,
								base.uiElement, TreeScope.Element, UIAeventHandler);
								
							try
							{
								lastSelectedState = this.IsSelected;
							}
							catch { }
						}
						else
						{
							UIAPropChangedEventHandler = new AutomationPropertyChangedEventHandler(
								OnUIAutomationPropChangedEvent);
							
							if (cfid == "WinForm")
							{
								Automation.AddAutomationPropertyChangedEventHandler(base.uiElement, TreeScope.Element,
									UIAPropChangedEventHandler, AutomationElement.NameProperty);
							}
							else
							{
								Automation.AddAutomationPropertyChangedEventHandler(base.uiElement, TreeScope.Element,
									UIAPropChangedEventHandler, SelectionItemPattern.IsSelectedProperty);
							}
						}
					}
					
					this.StateChangedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.StateChangedHandler -= value;
				
					if (this.StateChangedHandler == null)
					{
						string cfid = base.uiElement.Current.FrameworkId;
						if (cfid == "Win32")
						{
							RemoveEventHandlerWin32();
						}
						else
						{
							RemoveEventHandler();
						}
					}
				}
				catch {}
			}
		}
		
		private void RemoveEventHandlerWin32()
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

		private void RemoveEventHandler()
		{
			if (this.UIAPropChangedEventHandler == null)
			{
				return;
			}
			
			System.Threading.Tasks.Task.Run(() => 
			{
				try
				{
					Automation.RemoveAutomationPropertyChangedEventHandler(base.uiElement, 
						this.UIAPropChangedEventHandler);
					UIAPropChangedEventHandler = null;
				}
				catch { }
			}).Wait(5000);
		}
		
		private void OnUIAutomationPropChangedEvent(object sender, AutomationPropertyChangedEventArgs e)
		{
			if (e.Property.Id == SelectionItemPattern.IsSelectedProperty.Id && this.StateChangedHandler != null)
			{
				bool newValue = false;
				try
				{
					newValue = (bool)e.NewValue;
				}
				catch { }
				
				this.StateChangedHandler(this, newValue);
			}
			if (e.Property.Id == AutomationElement.NameProperty.Id && this.StateChangedHandler != null)
			{
				this.StateChangedHandler(this, this.IsSelected);
			}
		}
		
		//private DateTime lastChecked = DateTime.Now;
		private bool lastSelectedState = false;
		
		private void OnUIAutomationEvent(object sender, AutomationEventArgs e)
		{
			if (e.EventId == SelectionItemPattern.ElementSelectedEvent && this.StateChangedHandler != null)
			{
				/*DateTime currentChecked = DateTime.Now;
				if (currentChecked - lastChecked <= TimeSpan.FromMilliseconds(100))
				{
					lastChecked = currentChecked;
					return;
				}
				lastChecked = currentChecked;*/
			
				bool isSelected = false;
				try
				{
					isSelected = this.IsSelected;
				}
				catch { }
				
				if (isSelected != lastSelectedState)
				{
					lastSelectedState = isSelected;
					this.StateChangedHandler(this, isSelected);
				}
			}
		}
    }
}
