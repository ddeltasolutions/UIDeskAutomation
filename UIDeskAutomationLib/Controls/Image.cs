using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Represents an Image control.
    /// </summary>
    public class Image: ElementBase
    {
        public Image(AutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
