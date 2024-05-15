using System;
using System.Collections.Generic;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
	/// <summary>
    /// This class represents a header item.
    /// </summary>
    public class UIDA_HeaderItem : ElementBase
    {
        /// <summary>
        /// Creates a UDA_HeaderItem using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_HeaderItem(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets the Text of header item.
        /// </summary>
        public string Text
        {
            get
            {
                string textString = null;

                try
                {
                    textString = uiElement.Current.Name;
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("HeaderItem text: " + ex.Message);
                    throw new Exception("HeaderItem text: " + ex.Message);
                }

                return textString;
            }
        }
    }
}