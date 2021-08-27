using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Thumb control.
    /// </summary>
    public class UIDA_Thumb: ElementBase
    {
        public UIDA_Thumb(AutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
