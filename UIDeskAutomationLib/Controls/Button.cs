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
    public class Button: ElementBase
    {
		/// <summary>
        /// Creates a Button using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public Button(AutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
