using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Toolbar control.
    /// </summary>
    public class Toolbar: ElementBase
    {
        public Toolbar(AutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
