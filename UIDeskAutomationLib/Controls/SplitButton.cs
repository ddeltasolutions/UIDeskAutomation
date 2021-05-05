using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Represents a SplitButton control.
    /// </summary>
    public class SplitButton: ElementBase
    {
        public SplitButton(AutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
