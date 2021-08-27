using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Document UI element
    /// </summary>
    public class UIDA_Document: UIDA_Edit
    {
        /// <summary>
        /// Creates a UIDA_Document using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_Document(AutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
