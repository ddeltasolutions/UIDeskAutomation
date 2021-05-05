using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// This class represents a custom control.
    /// </summary>
    public class Custom: ElementBase
    {
        /// <summary>
        /// Creates a Custom using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public Custom(AutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
