using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using System.Threading;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a spinner UI control.
    /// </summary>
    public class UIDA_Spinner: GenericSpinner
    {
		/// <summary>
        /// Creates a UIDA_Spinner using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_Spinner(AutomationElement el): base(el) {}

        /// <summary>
        /// Increment the value of spinner. Is like pressing the up arrow.
        /// </summary>
        public void Increment()
        {
            Engine.TraceInLogFile("UIDA_Spinner.Increment method");
            UIDA_Button forwardButton = null;

            try
            {
                forwardButton = this.ButtonAt(null, 1/*, true*/);
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("UIDA_Spinner.Increment method failed");
                throw new Exception("UIDA_Spinner.Increment method failed");
            }

            if (forwardButton != null)
            {
                try
                {
                    forwardButton.Invoke();
                }
                catch 
                {
                    Engine.TraceInLogFile("UIDA_Spinner.Increment Invoke failed");
                    throw new Exception("UIDA_Spinner.Increment Invoke failed");
                }
            }
        }

        /// <summary>
        /// Decrements the value of spinner. Is like pressing the down arrow.
        /// </summary>
        public void Decrement()
        {
            UIDA_Button backwardButton = null;
            try
            {
                backwardButton = this.ButtonAt(null, 2/*, true*/);
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("UIDA_Spinner.Decrement method failed");
                throw new Exception("UIDA_Spinner.Decrement method failed");
            }

            if (backwardButton != null)
            {
                try
                {
                    backwardButton.Invoke();
                }
                catch
                {
                    Engine.TraceInLogFile("UIDA_Spinner.Decrement Invoke failed");
                    throw new Exception("UIDA_Spinner.Decrement Invoke failed");
                }
            }
        }
        
        /// <summary>
        /// Gets the minimum value the current spinner can get.
        /// </summary>
        /// <returns>The minimum value of the spinner</returns>
        new public double GetMinimum()
        {
            return base.GetMinimum();
        }
        
        /// <summary>
        /// Gets the maximum value the current spinner can get.
        /// </summary>
        /// <returns>The maximum value of the spinner</returns>
        new public double GetMaximum()
        {
            return base.GetMaximum();
        }
        
        /// <summary>
        /// Gets/Sets the value of the current spinner.
        /// </summary>
        new public double Value
        {
            get { return base.Value; }
            set { base.Value = value; }
        }
    }
}
