using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Class that represents a Title Bar ui element
    /// </summary>
    public class TitleBar: ElementBase
    {
        public TitleBar(AutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
