using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents an Image control.
    /// </summary>
    public class UIDA_Image: ElementBase
    {
        public UIDA_Image(AutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
