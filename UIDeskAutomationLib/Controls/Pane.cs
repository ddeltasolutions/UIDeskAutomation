using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Class that represents a Pane ui element
    /// </summary>
    public class Pane : ElementBase
    {
        public Pane(AutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
