using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Represents a Group UI element
    /// </summary>
    public class Group: ElementBase
    {
        /// <summary>
        /// Creates a Group using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public Group(AutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
