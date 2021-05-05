using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Represents a StatusBar control.
    /// </summary>
    public class StatusBar: ElementBase
    {
        public StatusBar(AutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
