using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a MenuBar ui element
    /// </summary>
    public class MenuBar: ElementBase
    {
        public MenuBar(AutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
