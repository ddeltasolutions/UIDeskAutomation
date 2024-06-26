﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using System.Threading;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a Menu Item ui element
    /// </summary>
    public class UIDA_MenuItem : ElementBase
    {
        public UIDA_MenuItem(AutomationElement el)
        {
            base.uiElement = el;
        }
        
        /// <summary>
        /// Accesses the menu item, like clicking on it
        /// </summary>
        public void AccessMenu()
        {
            base.Invoke();
        }

        /// <summary>
        /// Expands this menu item
        /// </summary>
        public void Expand()
        {
            if (this.IsAlive == false)
            {
                throw new Exception("This UI element is not available anymore.");
            }

            object objectPattern = null;
            this.uiElement.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out objectPattern);

            ExpandCollapsePattern expandCollapsePattern = objectPattern as ExpandCollapsePattern;
            if (expandCollapsePattern == null)
            {
                Engine.TraceInLogFile("Expand method - ExpandCollapse pattern not supported, Try Invoke");
                this.Invoke();
                return;
            }

            try
            {
                expandCollapsePattern.Expand();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Expand method - " + ex.Message);
                throw new Exception("Expand method - " + ex.Message);
            }
        }

        /// <summary>
        /// Collapses this menu item
        /// </summary>
        public void Collapse()
        {
            if (this.IsAlive == false)
            {
                throw new Exception("This UI element is not available anymore.");
            }

            object objectPattern = null;
            this.uiElement.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out objectPattern);
            ExpandCollapsePattern expandCollapsePattern = objectPattern as ExpandCollapsePattern;

            if (expandCollapsePattern == null)
            {
                Engine.TraceInLogFile("Collapse method - ExpandCollapse pattern not supported, Try Invoke");
                this.Invoke();
                return;
            }

            try
            {
                expandCollapsePattern.Collapse();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Collapse method - " + ex.Message);
                throw new Exception("Collapse method - " + ex.Message);
            }
        }

        /// <summary>
        /// Toggles a menu item (checked/unchecked)
        /// </summary>
        public void Toggle()
        {
            if (this.IsAlive == false)
            {
                throw new Exception("This UI element is not available to the user anymore.");
            }

            object objectPattern = null;

            this.uiElement.TryGetCurrentPattern(TogglePattern.Pattern, out objectPattern);
            TogglePattern togglePattern = objectPattern as TogglePattern;
            if (togglePattern == null)
            {
				this.Invoke();
				return;
            }

            try
            {
                togglePattern.Toggle();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Toggle method - " + ex.Message);
                throw new Exception("Toggle method - " + ex.Message);
            }
        }

        /// <summary>
        /// Gets or Sets the checked state of a menu item 
        /// </summary>
        public bool IsChecked
        {
            get
            {
                if (this.IsAlive == false)
                {
                    throw new Exception("This UI element is not available anymore.");
                }
				
				string frameworkid = "";
				try
				{
					frameworkid = this.uiElement.Current.FrameworkId;
				}
				catch {}

				if (frameworkid == "Win32")
				{
					object objectPattern = null;
					this.uiElement.TryGetCurrentPattern(TogglePattern.Pattern, out objectPattern);
					TogglePattern togglePattern = objectPattern as TogglePattern;

					if (togglePattern == null)
					{
						return false;
					}

					bool isChecked = false;
					try
					{
						if (togglePattern.Current.ToggleState == ToggleState.On)
						{
							isChecked = true;
						}
					}
					catch (Exception ex)
					{
						Engine.TraceInLogFile("Checked (get) property - " + ex.Message);
						throw new Exception("Checked (get) property - " + ex.Message);
					}
					
					return isChecked;
				}
				else
				{
					object objectPattern = null;
					this.uiElement.TryGetCurrentPattern(TogglePattern.Pattern, out objectPattern);
					TogglePattern togglePattern = objectPattern as TogglePattern;

					if (togglePattern == null)
					{
						Engine.TraceInLogFile("Checked property (get) - Toggle pattern not supported");
						throw new Exception("Checked property (get) - Toggle pattern not supported");
					}
					
					if (togglePattern.Current.ToggleState == ToggleState.On)
					{
						return true;
					}
					else
					{
						return false;
					}
				}
            }

            set
            {
                if (this.IsAlive == false)
                {
                    throw new Exception("This UI element is not available anymore.");
                }

				object objectPattern = null;
				
				this.uiElement.TryGetCurrentPattern(TogglePattern.Pattern, out objectPattern);
				TogglePattern togglePattern = objectPattern as TogglePattern;
				if (togglePattern == null)
				{
					string frameworkid = "";
					try
					{
						frameworkid = this.uiElement.Current.FrameworkId;
					}
					catch {}
					
					if (frameworkid == "Win32")
					{
						//current element is unchecked because it doesn't support TogglePattern
						if (value == true)
						{
							this.Invoke();
						}
						return;
					}
					else
					{
						Engine.TraceInLogFile("Checked (set) property - Toggle pattern not supported");
						throw new Exception("Checked (set) property - Toggle pattern not supported");
					}
				}

				try
				{
					/*if (((value == true) && (togglePattern.Current.ToggleState != ToggleState.On)) ||
						((value == false) && (togglePattern.Current.ToggleState == ToggleState.On)))*/
						
					if (((value == true) && (togglePattern.Current.ToggleState != ToggleState.On)) ||
						((value == false) && (togglePattern.Current.ToggleState != ToggleState.Off)))
					{
						togglePattern.Toggle();
					}
				}
				catch (Exception ex)
				{
					Engine.TraceInLogFile("Checked (set) property - " + ex.Message);
					throw new Exception("Checked (set) property - " + ex.Message);
				}
            }
        }
		
		/// <summary>
        /// Gets the text of the menu item
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
        /// Delegate for Expanded event
        /// </summary>
		/// <param name="sender">The menu item that sent the event.</param>
		public delegate void Expanded(UIDA_MenuItem sender);
		internal Expanded ExpandedHandler = null;
		
		/// <summary>
        /// Attaches/detaches a handler to expanded event
        /// </summary>
		public event Expanded ExpandedEvent
		{
			add
			{
				try
				{
					if (this.ExpandedHandler == null)
					{
						if (base.uiElement.Current.FrameworkId == "WinForm")
						{
							UIAeventHandler = new AutomationEventHandler(OnUIAutomationEvent);
						
							Automation.AddAutomationEventHandler(AutomationElement.MenuOpenedEvent,
								base.uiElement, TreeScope.Subtree, UIAeventHandler);
						}
						else
						{
							UIAPropChangedEventHandler = new AutomationPropertyChangedEventHandler(
								OnUIAutomationPropChangedEvent);

							Automation.AddAutomationPropertyChangedEventHandler(base.uiElement, TreeScope.Element,
									UIAPropChangedEventHandler, ExpandCollapsePattern.ExpandCollapseStateProperty);
						}
					}
					
					this.ExpandedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.ExpandedHandler -= value;
				
					if (this.ExpandedHandler == null)
					{
						if (base.uiElement.Current.FrameworkId == "WinForm")
						{
							RemoveEventHandlerWinForm();
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
		
		private void RemoveEventHandlerWinForm()
		{
			if (this.UIAeventHandler == null)
			{
				return;
			}
			
			System.Threading.Tasks.Task.Run(() => 
			{
				try
				{
					Automation.RemoveAutomationEventHandler(AutomationElement.MenuOpenedEvent, 
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
			if (e.Property.Id == ExpandCollapsePattern.ExpandCollapseStateProperty.Id && this.ExpandedHandler != null)
			{
				try
				{
					ExpandCollapseState state = (ExpandCollapseState)e.NewValue;
					if (state == ExpandCollapseState.Expanded)
					{
						ExpandedHandler(this);
					}
				}
				catch { }
			}
		}
		
		private AutomationEventHandler UIA_ClickedEventHandler = null;
		
		/// <summary>
        /// Delegate for Clicked event
        /// </summary>
		/// <param name="sender">The menu item that sent the event.</param>
		public delegate void Clicked(UIDA_MenuItem sender);
		internal Clicked ClickedHandler = null;
		
		/// <summary>
        /// Attaches/detaches a handler to click event
        /// </summary>
		public event Clicked ClickedEvent
		{
			add
			{
				try
				{
					if (this.ClickedHandler == null)
					{
						this.UIA_ClickedEventHandler = new AutomationEventHandler(OnUIAutomationEvent);
		
						Automation.AddAutomationEventHandler(InvokePattern.InvokedEvent,
							base.uiElement, TreeScope.Element, UIA_ClickedEventHandler);
					}
					
					this.ClickedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.ClickedHandler -= value;
				
					if (this.ClickedHandler == null)
					{
						if (this.UIA_ClickedEventHandler == null)
						{
							return;
						}
						
						System.Threading.Tasks.Task.Run(() => 
						{
							try
							{
								Automation.RemoveAutomationEventHandler(InvokePattern.InvokedEvent, 
									base.uiElement, this.UIA_ClickedEventHandler);
								UIA_ClickedEventHandler = null;
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
			if (e.EventId == InvokePattern.InvokedEvent && this.ClickedHandler != null)
			{
				ClickedHandler(this);
			}
			else if (e.EventId == AutomationElement.MenuOpenedEvent && this.ExpandedHandler != null)
			{
				ExpandedHandler(this);
			}
		}
    }
}
