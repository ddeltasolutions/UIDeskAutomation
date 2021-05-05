using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Represents a Thumb control.
    /// </summary>
    public class Thumb: ElementBase
    {
        public Thumb(AutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
