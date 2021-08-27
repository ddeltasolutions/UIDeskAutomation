using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a Title Bar ui element
    /// </summary>
    public class UIDA_TitleBar: ElementBase
    {
        public UIDA_TitleBar(AutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
