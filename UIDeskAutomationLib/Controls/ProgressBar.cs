using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Represents the class for a Progress Bar UI control.
    /// </summary>
    public class ProgressBar: GenericSpinner
    {
        public ProgressBar(AutomationElement el): base(el)
        { }
        
        /// <summary>
        /// Gets the value of the current progressbar.
        /// </summary>
        public double Value
        {
            get { return base.Value; }
        }
        
        /// <summary>
        /// Gets the minimum value the current progressbar can get.
        /// </summary>
        /// <returns>The minimum value of the progressbar</returns>
        public double GetMinimum()
        {
            return base.GetMinimum();
        }
        
        /// <summary>
        /// Gets the maximum value the current progressbar can get.
        /// </summary>
        /// <returns>The maximum value of the progressbar</returns>
        public double GetMaximum()
        {
            return base.GetMaximum();
        }
    }
}
