using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Represents a Tooltip control.
    /// </summary>
    public class Tooltip: ElementBase
    {
        public Tooltip(AutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
