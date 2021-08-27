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
    public class UIDA_MenuBar: ElementBase
    {
        public UIDA_MenuBar(AutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
