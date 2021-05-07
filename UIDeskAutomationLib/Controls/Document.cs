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
    public class Document: Edit
    {
        /// <summary>
        /// Creates a Document using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public Document(AutomationElement el)
        {
            base.uiElement = el;
        }
    }
}
