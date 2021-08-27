using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a SplitButton control.
    /// </summary>
    public class UIDA_SplitButton: ElementBase
    {
        public UIDA_SplitButton(AutomationElement el)
        {
            this.uiElement = el;
        }
        
        /// <summary>
        /// Presses the split button
        /// </summary>
        public void Press()
        {
            base.Invoke();
        }
    }
}
