using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// This class represents a Static Text control.
    /// </summary>
    public class Label: ElementBase
    {
        public Label(AutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
