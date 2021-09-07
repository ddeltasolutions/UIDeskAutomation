using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a button UI element
    /// </summary>
    public class UIDA_Button: ElementBase
    {
		/// <summary>
        /// Creates a UIDA_Button using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_Button(AutomationElement el)
        {
            base.uiElement = el;
        }
        
        /// <summary>
        /// Presses the button
        /// </summary>
        public void Press()
        {
            base.Invoke();
        }
		
		/// <summary>
        /// Gets a boolean to determine if the button is pressed or not. This is supported only for toggle buttons.
        /// </summary>
        public bool IsPressed
        {
            get
            {
                if (this.IsAlive == false)
                {
                    throw new Exception("This UI element is not available to the user anymore.");
                }

                object togglePatternObject = null;

                if (base.uiElement.TryGetCurrentPattern(TogglePattern.Pattern,
                    out togglePatternObject) == true)
                {
                    TogglePattern togglePattern = togglePatternObject as TogglePattern;
                    if (togglePattern != null)
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
                }

                //return null;
				throw new Exception("Could not get the pressed state of the button. This is supported only for toggle buttons.");
            }
            /*set
            {
				if (this.IsAlive == false)
                {
                    throw new Exception("This UI element is not available to the user anymore.");
                }

                object togglePatternObject = null;

                if (base.uiElement.TryGetCurrentPattern(TogglePattern.Pattern,
                    out togglePatternObject) == true)
                {
                    TogglePattern togglePattern = togglePatternObject as TogglePattern;
                    if (togglePattern != null)
                    {
                        if ((value == false && togglePattern.Current.ToggleState == ToggleState.On) ||
							(value == true && togglePattern.Current.ToggleState != ToggleState.On))
                        {
							togglePattern.Toggle();
                        }
						return;
                    }
                }
				
				throw new Exception("Could not set the pressed state of the button. This is supported only for toggle buttons.");
            }*/
        }
    }
}
