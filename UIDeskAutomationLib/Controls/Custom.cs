using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// This class represents a custom control.
    /// </summary>
    public class UIDA_Custom: ElementBase
    {
        /// <summary>
        /// Creates a UIDA_Custom using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_Custom(AutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
