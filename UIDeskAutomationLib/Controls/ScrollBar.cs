using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a ScrollBar control.
    /// </summary>
    public class UIDA_ScrollBar: GenericSpinner
    {
		/// <summary>
        /// Creates a UIDA_ScrollBar using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_ScrollBar(AutomationElement el): base(el) {}

        /// <summary>
        /// Increments the value of ScrollBar. 
        /// Is like pressing the right button/arrow (for horizontal scrollbar) or 
        /// the bottom button/down arrow (for vertical scrollbar) of the ScrollBar control.
        /// </summary>
        public void SmallIncrement()
        {
            double smallChange = 0.0;

            try
            {
                smallChange = base.GetSmallChange();
				double maximum = base.GetMaximum();
				if (base.Value + smallChange <= maximum)
				{
					base.Value += smallChange;
				}
				else
				{
					base.Value = maximum;
				}
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("ScrollBar.SmallIncrement: " + ex.Message);
				throw ex;
            }
        }

        /// <summary>
        /// Increments scrollbar control with a larger value.
        /// Is like pressing the region between the thumb and the right arrow at 
        /// a horizontal scrollbar or pressing the region between the thumb and the 
        /// down arrow for a vertical scrollbar.
        /// </summary>
        public void LargeIncrement()
        {
            try
            {
                double largeChange = base.GetLargeChange();
				double maximum = base.GetMaximum();
				if (base.Value + largeChange <= maximum)
				{
					base.Value += largeChange;
				}
				else
				{
					base.Value = maximum;
				}
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("ScrollBar.LargeIncrement: " + ex.Message);
				throw ex;
            }
        }

        /// <summary>
        /// Decrements the value of ScrollBar. 
        /// Is like pressing the left button (arrow) for a horizontal scrollbar or 
        /// the top button (up arrow) for a vertical scrollbar.
        /// </summary>
        public void SmallDecrement()
        {
            try
            {
                double smallChange = base.GetSmallChange();
				double minimum = base.GetMinimum();
				if (base.Value - smallChange >= minimum)
				{
					base.Value -= smallChange;
				}
				else
				{
					base.Value = minimum;
				}
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("ScrollBar.SmallDecrement: " + ex.Message);
				throw ex;
            }
        }

        /// <summary>
        /// Decrements the value of a scrollbar with a larger value.
        /// Is like pressing the region between left arrow and the thumb for a 
        /// horizontal scrollbar or like pressing the region between up arrow and the 
        /// thumb for a vertical scrollbar.
        /// </summary>
        public void LargeDecrement()
        {
            try
            {
                double largeChange = base.GetLargeChange();
				double minimum = base.GetMinimum();
				if (base.Value - largeChange >= minimum)
				{
					base.Value -= largeChange;
				}
				else
				{
					base.Value = minimum;
				}
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("ScrollBar.LargeDecrement: " + ex.Message);
				throw ex;
            }
        }

        /// <summary>
        /// Gets the minimum value the current scrollbar can get.
        /// </summary>
        /// <returns>The minimum value of the scrollbar</returns>
        new public double GetMinimum()
        {
            return base.GetMinimum();
        }
        
        /// <summary>
        /// Gets the maximum value the current scrollbar can get.
        /// </summary>
        /// <returns>The maximum value of the scrollbar</returns>
        new public double GetMaximum()
        {
            return base.GetMaximum();
        }
        
        /// <summary>
        /// Gets/Sets the value of the current scrollbar.
        /// </summary>
        new public double Value
        {
            get { return base.Value; }
            set { base.Value = value; }
        }
		
		/// <summary>
        /// Attaches/detaches a handler to value changed event. You can cast the first parameter (sender - of type GenericSpinner) to an UIDA_ScrollBar object.
		/// The second parameter (of type double) is the new value of the scroll bar.
        /// </summary>
		public event ValueChanged ValueChangedEvent
		{
			add
			{
				base.ValueChangedEvent += value;
			}
			remove
			{
				base.ValueChangedEvent -= value;
			}
		}
    }
}
