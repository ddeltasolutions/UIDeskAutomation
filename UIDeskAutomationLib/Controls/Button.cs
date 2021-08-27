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
    }
}
