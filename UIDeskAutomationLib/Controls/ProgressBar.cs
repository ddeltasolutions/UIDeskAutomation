using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents the class for a Progress Bar UI control.
    /// </summary>
    public class UIDA_ProgressBar: GenericSpinner
    {
		/// <summary>
        /// Creates a UIDA_ProgressBar using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_ProgressBar(AutomationElement el): base(el)
        { }
        
        /// <summary>
        /// Gets the value of the current progressbar.
        /// </summary>
        new public double Value
        {
            get { return base.Value; }
        }
        
        /// <summary>
        /// Gets the minimum value the current progressbar can get.
        /// </summary>
        /// <returns>The minimum value of the progressbar</returns>
        new public double GetMinimum()
        {
            return base.GetMinimum();
        }
        
        /// <summary>
        /// Gets the maximum value the current progressbar can get.
        /// </summary>
        /// <returns>The maximum value of the progressbar</returns>
        new public double GetMaximum()
        {
            return base.GetMaximum();
        }
    }
}
