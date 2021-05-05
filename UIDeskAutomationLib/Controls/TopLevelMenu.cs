using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Class that represents a Top Level Menu element, like a Contextual menu.
    /// </summary>
    public class TopLevelMenu : ElementBase
    {
        public TopLevelMenu(AutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
