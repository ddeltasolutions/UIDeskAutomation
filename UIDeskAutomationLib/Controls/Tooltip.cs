using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Tooltip control.
    /// </summary>
    public class UIDA_Tooltip: ElementBase
    {
        public UIDA_Tooltip(AutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
