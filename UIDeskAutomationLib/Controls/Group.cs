using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Group UI element
    /// </summary>
    public class UIDA_Group: ElementBase
    {
        /// <summary>
        /// Creates a UIDA_Group using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_Group(AutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
