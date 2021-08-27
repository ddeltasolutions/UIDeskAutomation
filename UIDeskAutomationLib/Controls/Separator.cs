using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Separator control.
    /// </summary>
    public class UIDA_Separator: ElementBase
    {
        public UIDA_Separator(AutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
