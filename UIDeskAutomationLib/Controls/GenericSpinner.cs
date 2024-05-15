using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Generic base class for Spinner, Slider, ProgressBar and Scrollbar.
    /// This class is used only internally and cannot be instantiated.
    /// </summary>
    public abstract class GenericSpinner: ElementBase
    {
        internal GenericSpinner(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets/Sets the value of the current element spinner/slider/progressbar.
        /// Value of a progress bar cannot be set, only read.
        /// </summary>
        internal double Value
        {
            get
            {
                string localizedControlType = "current element";

                try
                {
                    localizedControlType = this.uiElement.Current.LocalizedControlType;
                }
                catch { }

                try
                {
                    if (this.uiElement.Current.ControlType == ControlType.Spinner)
                    {
                        double? val = GetSetWindowsFormsSpinnerValue();

                        if (val.HasValue)
                        {
                            return val.Value;
                        }
                    }

                }
                catch { }

                object rangeValuePatternObj = null;
				this.uiElement.TryGetCurrentPattern(RangeValuePattern.Pattern, out rangeValuePatternObj);
				RangeValuePattern rangeValuePattern = rangeValuePatternObj as RangeValuePattern;

				if (rangeValuePattern == null)
				{
					object valuePatternObj = null;
					this.uiElement.TryGetCurrentPattern(ValuePattern.Pattern, out valuePatternObj);
					ValuePattern valuePattern = valuePatternObj as ValuePattern;
				
					if (valuePattern == null)
					{
						Engine.TraceInLogFile("Cannot get value of " + localizedControlType);
						throw new Exception("Cannot get value of " + localizedControlType);
					}
					
					try
					{
						return double.Parse(valuePattern.Current.Value);
					}
					catch (Exception ex)
					{
						Engine.TraceInLogFile("Cannot get " + localizedControlType + " value: " + ex.Message);
						throw new Exception("Cannot get " + localizedControlType + " value: " + ex.Message);
					}
				}

				try
				{
					return rangeValuePattern.Current.Value;
				}
				catch (Exception ex)
				{
					Engine.TraceInLogFile("Cannot get " + localizedControlType + " value: " + ex.Message);
					throw new Exception("Cannot get " + localizedControlType + " value: " + ex.Message);
				}
            }

            set
            {
                try
                {
                    // Don't let the user set the value of progress bar control.
                    if (this.uiElement.Current.ControlType == ControlType.ProgressBar)
                    {
                        Engine.TraceInLogFile("Cannot set the value of a progress bar");
                        throw new Exception("Cannot set the value of a progress bar");
                    }
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("Set value error: " + ex.Message);
                    throw new Exception("Set value error: " + ex.Message);
                }

                string localizedControlType = "current element";
                try
                {
                    localizedControlType = this.uiElement.Current.LocalizedControlType;
                }
                catch { }

                object rangeValuePatternObj = null;
				string message = null;

                if (this.uiElement.TryGetCurrentPattern(RangeValuePattern.Pattern,
                    out rangeValuePatternObj) == true)
                {
                    RangeValuePattern rangeValuePattern = rangeValuePatternObj as
                        RangeValuePattern;

					if (rangeValuePattern != null)
					{
						try
						{
							double maximum = rangeValuePattern.Current.Maximum;
							if (value > maximum)
							{
								value = maximum;
							}
							rangeValuePattern.SetValue(value);
							return;
						}
						catch (Exception ex)
						{
							message = "Cannot set value of " + localizedControlType + ": " + ex.Message;
						}
					}
                }
				
				try
                {
                    if (this.uiElement.Current.ControlType == ControlType.Spinner)
                    {
                        if (GetSetWindowsFormsSpinnerValue(value, true) != null)
						{
							return;
						}
                    }
                }
                catch { }

				if (message != null)
				{
					Engine.TraceInLogFile(message);
					throw new Exception(message);
				}
				else
				{
					Engine.TraceInLogFile("Cannot set value of " + localizedControlType);
					throw new Exception("Cannot set value of " + localizedControlType);
				}
            }
        }

        // for set = false, 'value' parameter is ignored
        private double? GetSetWindowsFormsSpinnerValue(double value = 0.0, 
            bool set = false)
        {
            IntPtr windowHandle = this.GetWindow();

            if (windowHandle == IntPtr.Zero)
            {
                return null;
            }

            StringBuilder className = new StringBuilder(256);
            UnsafeNativeFunctions.GetClassName(windowHandle, className, 256);

            if (className.ToString().StartsWith("WindowsForms"))
            {
                // Windows Forms spinner
                IntPtr hwndEdit = IntPtr.Zero;

                IntPtr hwndChild = UnsafeNativeFunctions.FindWindowEx(
                    windowHandle, IntPtr.Zero, null, null);

                while (hwndChild != IntPtr.Zero)
                {
                    StringBuilder childClassName = new StringBuilder(256);
                    UnsafeNativeFunctions.GetClassName(hwndChild, childClassName, 256);
                    if (childClassName.ToString().ToLower().Contains("edit"))
                    {
                        hwndEdit = hwndChild;
                        break;
                    }

                    hwndChild = UnsafeNativeFunctions.FindWindowEx(windowHandle,
                        hwndChild, null, null);
                }

                if (hwndEdit != IntPtr.Zero)
                {
                    if (set == false)
                    {
                        //get window text
                        IntPtr textLengthPtr = UnsafeNativeFunctions.SendMessage(hwndEdit,
                        WindowMessages.WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero);

                        string windowText = string.Empty;

                        if (textLengthPtr.ToInt32() > 0)
                        {
                            int textLength = textLengthPtr.ToInt32() + 1;
                            StringBuilder text = new StringBuilder(textLength);

                            UnsafeNativeFunctions.SendMessage(hwndEdit,
                                WindowMessages.WM_GETTEXT, textLength, text);

                            windowText = text.ToString();
                        }

                        double valueDouble = 0.0;
                        if (double.TryParse(windowText, out valueDouble) == true)
                        {
                            return valueDouble;
                        }
                    }
                    else
                    { 
                        // set spinner value
                        string valueString = string.Empty;
                        try
                        {
                            valueString = Convert.ToInt32(value).ToString();
                        }
                        catch 
                        {
                            return null;
                        }

                        IntPtr textPtr = Marshal.StringToBSTR(valueString);
                        if (textPtr != IntPtr.Zero)
                        {
                            UnsafeNativeFunctions.SendMessage(hwndEdit,
                                WindowMessages.WM_SETTEXT, IntPtr.Zero, textPtr);
							return 0;
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Gets the minimum value the current element spinner/slider/progressbar can get.
        /// </summary>
        /// <returns></returns>
        internal double GetMinimum()
        {
            string localizedControlType = "current element";

            try
            {
                localizedControlType = this.uiElement.Current.LocalizedControlType;
            }
            catch { }

            object rangeValuePatternObj = null;

            if (this.uiElement.TryGetCurrentPattern(RangeValuePattern.Pattern,
                out rangeValuePatternObj) == true)
            {
                RangeValuePattern rangeValuePattern = rangeValuePatternObj as RangeValuePattern;

                if (rangeValuePattern == null)
                {
                    Engine.TraceInLogFile("Cannot get minimum of " + localizedControlType);
                    throw new Exception("Cannot get minimum of " + localizedControlType);
                }

                try
                {
                    return rangeValuePattern.Current.Minimum;
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("Cannot get minimum of " + localizedControlType + ": " + ex.Message);
                    throw new Exception("Cannot get minimum of " + localizedControlType + ": " + ex.Message);
                }
            }

            Engine.TraceInLogFile("Cannot get minimum of " + localizedControlType);
            throw new Exception("Cannot get minimum of " + localizedControlType);
        }

        /// <summary>
        /// Gets the maximum value the current element spinner/slider/progressbar can get.
        /// </summary>
        /// <returns></returns>
        internal double GetMaximum()
        {
            string localizedControlType = "current element";

            try
            {
                localizedControlType = this.uiElement.Current.LocalizedControlType;
            }
            catch { }

            object rangeValuePatternObj = null;

            if (this.uiElement.TryGetCurrentPattern(RangeValuePattern.Pattern,
                out rangeValuePatternObj) == true)
            {
                RangeValuePattern rangeValuePattern = rangeValuePatternObj as RangeValuePattern;

                if (rangeValuePattern == null)
                {
                    Engine.TraceInLogFile("Cannot get maximum of " + localizedControlType);
                    throw new Exception("Cannot get maximum of " + localizedControlType);
                }

                try
                {
                    return rangeValuePattern.Current.Maximum;
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("Cannot get maximum of " + localizedControlType + ": " + ex.Message);
                    throw new Exception("Cannot get maximum of " + localizedControlType + ": " + ex.Message);
                }
            }

            Engine.TraceInLogFile("Cannot get maximum of " + localizedControlType);
            throw new Exception("Cannot get maximum of " + localizedControlType);
        }

        internal double GetSmallChange()
        {
            object rangeValuePatternObj = null;
            this.uiElement.TryGetCurrentPattern(RangeValuePattern.Pattern, out rangeValuePatternObj);

            RangeValuePattern rangeValuePattern = rangeValuePatternObj as RangeValuePattern;
            if (rangeValuePattern == null)
            {
                Engine.TraceInLogFile("GetSmallChange() method: RangeValuePattern not supported");
                throw new Exception("GetSmallChange() method: RangeValuePattern not supported");
            }

            try
            {
                double smallChange = rangeValuePattern.Current.SmallChange;
                return smallChange;
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("GetSmallChange() method failed: " + ex.Message);
                throw new Exception("GetSmallChange() method failed: " + ex.Message);
            }
        }

        internal double GetLargeChange()
        {
            object rangeValuePatternObj = null;
            this.uiElement.TryGetCurrentPattern(RangeValuePattern.Pattern, out rangeValuePatternObj);

            RangeValuePattern rangeValuePattern = rangeValuePatternObj as RangeValuePattern;
            if (rangeValuePattern == null)
            {
                Engine.TraceInLogFile("GetLargeChange() method: RangeValuePattern not supported");
                throw new Exception("GetLargeChange() method: RangeValuePattern not supported");
            }

            try
            {
                double largeChange = rangeValuePattern.Current.LargeChange;
                return largeChange;
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("GetLargeChange() method failed: " + ex.Message);
                throw new Exception("GetLargeChange() method failed: " + ex.Message);
            }
        }
		
		private AutomationPropertyChangedEventHandler UIA_ValueChangedEventHandler = null;
		
		public delegate void ValueChanged(GenericSpinner sender, double newValue);
		internal ValueChanged ValueChangedHandler = null;
		
		/// <summary>
        /// Attaches/detaches a handler to value changed event
        /// </summary>
		internal event ValueChanged ValueChangedEvent
		{
			add
			{
				try
				{
					if (this.ValueChangedHandler == null)
					{
						this.UIA_ValueChangedEventHandler = new AutomationPropertyChangedEventHandler(OnUIAutomationPropChangedEvent);
			
						if (base.uiElement.Current.FrameworkId == "WinForm")
						{
							Automation.AddAutomationPropertyChangedEventHandler(base.uiElement, TreeScope.Element,
								UIA_ValueChangedEventHandler, ValuePattern.ValueProperty);
						}
						else
						{
							Automation.AddAutomationPropertyChangedEventHandler(base.uiElement, TreeScope.Element,
								UIA_ValueChangedEventHandler, RangeValuePattern.ValueProperty);
						}
					}
					
					this.ValueChangedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.ValueChangedHandler -= value;
				
					if (this.ValueChangedHandler == null)
					{
						if (this.UIA_ValueChangedEventHandler == null)
						{
							return;
						}
						
						System.Threading.Tasks.Task.Run(() => 
						{
							try
							{
								Automation.RemoveAutomationPropertyChangedEventHandler(base.uiElement, 
									this.UIA_ValueChangedEventHandler);
								UIA_ValueChangedEventHandler = null;
							}
							catch { }
						}).Wait(5000);
					}
				}
				catch {}
			}
		}
		
		private void OnUIAutomationPropChangedEvent(object sender, AutomationPropertyChangedEventArgs e)
		{
			if (e.Property.Id == RangeValuePattern.ValueProperty.Id && this.ValueChangedHandler != null)
			{
				double newValue = 0.0;
				try
				{
					newValue = (double)e.NewValue;
				}
				catch { }
				
				ValueChangedHandler(this, newValue);
			}
			if (e.Property.Id == ValuePattern.ValueProperty.Id && this.ValueChangedHandler != null)
			{
				ValueChangedHandler(this, this.Value);
			}
		}
    }
}
