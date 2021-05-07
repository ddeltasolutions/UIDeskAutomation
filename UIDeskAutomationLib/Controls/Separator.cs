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
    public class Separator: ElementBase
    {
        public Separator(AutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
