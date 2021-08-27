using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a Pane ui element
    /// </summary>
    public class UIDA_Pane : ElementBase
    {
        public UIDA_Pane(AutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
