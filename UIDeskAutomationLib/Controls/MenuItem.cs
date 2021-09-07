using System;
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
    }
}
