using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Class that represents a Slider UI element.
    /// </summary>
    public class Slider: GenericSpinner
    {
        public Slider(AutomationElement el)
            : base(el)
        { }

        /// <summary>
        /// Increments the value of Slider. Is like pressing in the right side of the thumb.
        /// </summary>
        public void Increment()
        {
            Button increaseButton = null;

            try
            {
                increaseButton = this.ButtonAt(null, 2);
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Slider.Increment method failed");
                throw new Exception("Slider.Increment method failed");
            }

            if (increaseButton != null)
            {
                increaseButton.Invoke();
            }
        }

        /// <summary>
        /// Decrements the value of Slider. Is like pressing in the left side of the thumb.
        /// </summary>
        public void Decrement()
        {
            Button decreaseButton = null;
            try
            {
                decreaseButton = this.ButtonAt(null, 1);
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Slider.Decrement method failed");
                throw new Exception("Slider.Decrement method failed");
            }

            if (decreaseButton != null)
            {
                decreaseButton.Invoke();
            }
        }
        
        /// <summary>
        /// Gets the minimum value the current slider can get.
        /// </summary>
        /// <returns>The minimum value of the slider</returns>
        public double GetMinimum()
        {
            return base.GetMinimum();
        }
        
        /// <summary>
        /// Gets the maximum value the current slider can get.
        /// </summary>
        /// <returns>The maximum value of the slider</returns>
        public double GetMaximum()
        {
            return base.GetMaximum();
        }
        
        /// <summary>
        /// Gets/Sets the value of the current slider.
        /// </summary>
        public double Value
        {
            get { return base.Value; }
            set { base.Value = value; }
        }
    }
}
