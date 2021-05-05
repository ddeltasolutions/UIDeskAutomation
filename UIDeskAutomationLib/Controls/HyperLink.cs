using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Represents a Hyper Link control.
    /// </summary>
    public class HyperLink: ElementBase
    {
        public HyperLink(AutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
