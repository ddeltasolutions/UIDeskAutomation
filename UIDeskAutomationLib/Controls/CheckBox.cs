using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a CheckBox UI element
    /// </summary>
    public class UIDA_CheckBox : ElementBase
    {
        /// <summary>
        /// Creates a CheckBoxElement using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_CheckBox(AutomationElement el)
        {
            base.uiElement = el;
        }

        /// <summary>
        /// Checks a checkbox
        /// </summary>
        public void Check()
        {
            if (this.IsAlive == false)
            {
                throw new Exception("This UI element is not available to the user anymore.");
            }

            object togglePatternObject = null;

            if (base.uiElement.TryGetCurrentPattern(TogglePattern.Pattern, 
                out togglePatternObject) == true)
            {
                TogglePattern togglePattern = togglePatternObject as TogglePattern;
                Debug.Assert(togglePattern != null);

                if (togglePattern.Current.ToggleState != ToggleState.On)
                {
					if (this.uiElement.Current.FrameworkId == "WPF")
					{
						// for WPF use mouse click because TogglePattern.Toggle() does not call the checkbox's Click event handler
						this.Click();
						Engine.GetInstance().Sleep(100);
					}
					else
					{
						togglePattern.Toggle();
					}
                }
				
				if (togglePattern.Current.ToggleState != ToggleState.On)
                {
					if (this.uiElement.Current.FrameworkId == "WPF")
					{
						this.Click();
						Engine.GetInstance().Sleep(100);
					}
					else
					{
						togglePattern.Toggle();
					}
                }

				if (togglePattern.Current.ToggleState == ToggleState.On)
				{
					return;
				}
            }

            // current element does not support Toggle Pattern
            // todo: simulate click
            IntPtr hwnd = IntPtr.Zero;

            try
            {
                hwnd = new IntPtr(this.uiElement.Current.NativeWindowHandle);
            }
            catch
            { }

            bool isWin32Button = false;

            if (hwnd != IntPtr.Zero)
            {
                StringBuilder className = new StringBuilder(256);
                UnsafeNativeFunctions.GetClassName(hwnd, className, 256);

                if (className.ToString() == "Button")
                {
                    isWin32Button = true;

                    // common Win32 checkbox window
                    IntPtr result = UnsafeNativeFunctions.SendMessage(hwnd,
                        ButtonMessages.BM_GETCHECK, IntPtr.Zero, IntPtr.Zero);

                    if (result.ToInt32() != (int)ButtonMessages.BST_CHECKED)
                    {
                        UnsafeNativeFunctions.SendMessage(hwnd, ButtonMessages.BM_SETCHECK,
                            new IntPtr(ButtonMessages.BST_CHECKED), IntPtr.Zero);
                    }
                }
            }

            if (!isWin32Button)
            {

            }
        }

        /// <summary>
        /// Unchecks a checkbox
        /// </summary>
        public void Uncheck()
        {
            if (this.IsAlive == false)
            {
                throw new Exception("This UI element is not available to the user anymore.");
            }

            object togglePatternObject = null;

            if (base.uiElement.TryGetCurrentPattern(TogglePattern.Pattern,
                out togglePatternObject) == true)
            {
                //object supports Toggle pattern
                TogglePattern togglePattern = togglePatternObject as TogglePattern;
                Debug.Assert(togglePattern != null);

                if (togglePattern.Current.ToggleState != ToggleState.Off)
                {
					if (this.uiElement.Current.FrameworkId == "WPF")
					{
						this.Click();
						Engine.GetInstance().Sleep(100);
					}
					else
					{
						togglePattern.Toggle();
					}
                }
				
				if (togglePattern.Current.ToggleState != ToggleState.Off)
                {
					if (this.uiElement.Current.FrameworkId == "WPF")
					{
						this.Click();
						Engine.GetInstance().Sleep(100);
					}
					else
					{
						togglePattern.Toggle();
					}
                }

				if (togglePattern.Current.ToggleState == ToggleState.Off)
				{
					return;
				}
            }

            // current element does not support Toggle Pattern
            // use Win32 messages for common Win32 checkboxes
            IntPtr hwnd = IntPtr.Zero;

            try
            {
                hwnd = new IntPtr(this.uiElement.Current.NativeWindowHandle);
            }
            catch
            { }

            bool isWin32Button = false;

            if (hwnd != IntPtr.Zero)
            {
                StringBuilder className = new StringBuilder(256);
                UnsafeNativeFunctions.GetClassName(hwnd, className, 256);

                if (className.ToString() == "Button")
                {
                    isWin32Button = true;

                    // common Win32 checkbox window
                    IntPtr result = UnsafeNativeFunctions.SendMessage(hwnd,
                        ButtonMessages.BM_GETCHECK, IntPtr.Zero, IntPtr.Zero);

                    if (result.ToInt32() != (int)ButtonMessages.BST_UNCHECKED)
                    {
                        UnsafeNativeFunctions.SendMessage(hwnd, ButtonMessages.BM_SETCHECK,
                            new IntPtr(ButtonMessages.BST_UNCHECKED), IntPtr.Zero);
                    }
                }
            }

            if (!isWin32Button)
            {
                //handle when checkbox is not a common Win32 checkbox
            }
        }
		
		/// <summary>
        /// Puts the checkbox in an indeterminate state if it is possible
        /// </summary>
        public void SetIndeterminate()
        {
            if (this.IsAlive == false)
            {
                throw new Exception("This UI element is not available to the user anymore.");
            }

            object togglePatternObject = null;

            if (base.uiElement.TryGetCurrentPattern(TogglePattern.Pattern, 
                out togglePatternObject) == true)
            {
                TogglePattern togglePattern = togglePatternObject as TogglePattern;
                Debug.Assert(togglePattern != null);

                if (togglePattern.Current.ToggleState != ToggleState.Indeterminate)
                {
					if (this.uiElement.Current.FrameworkId == "WPF")
					{
						this.Click();
						Engine.GetInstance().Sleep(100);
					}
					else
					{
						togglePattern.Toggle();
					}
                }
				
				if (togglePattern.Current.ToggleState != ToggleState.Indeterminate)
                {
					if (this.uiElement.Current.FrameworkId == "WPF")
					{
						this.Click();
						Engine.GetInstance().Sleep(100);
					}
					else
					{
						togglePattern.Toggle();
					}
                }

				if (togglePattern.Current.ToggleState == ToggleState.Indeterminate)
				{
					return;
				}
            }

            // current element does not support Toggle Pattern
            // todo: simulate click
            IntPtr hwnd = IntPtr.Zero;

            try
            {
                hwnd = new IntPtr(this.uiElement.Current.NativeWindowHandle);
            }
            catch
            { }

            bool isWin32Button = false;

            if (hwnd != IntPtr.Zero)
            {
                StringBuilder className = new StringBuilder(256);
                UnsafeNativeFunctions.GetClassName(hwnd, className, 256);

                if (className.ToString() == "Button")
                {
                    isWin32Button = true;

                    // common Win32 checkbox window
                    IntPtr result = UnsafeNativeFunctions.SendMessage(hwnd,
                        ButtonMessages.BM_GETCHECK, IntPtr.Zero, IntPtr.Zero);

                    if (result.ToInt32() != (int)ButtonMessages.BST_INDETERMINATE)
                    {
                        UnsafeNativeFunctions.SendMessage(hwnd, ButtonMessages.BM_SETCHECK,
                            new IntPtr(ButtonMessages.BST_INDETERMINATE), IntPtr.Zero);
                    }
                }
            }

            if (!isWin32Button)
            {

            }
        }
		
		/// <summary>
        /// Toggles between the states of a checkbox
        /// </summary>
        public void Toggle()
        {
            if (this.IsAlive == false)
            {
                throw new Exception("This UI element is not available to the user anymore.");
            }
			
			if (this.uiElement.Current.FrameworkId == "WPF")
			{
				this.Click();
				Engine.GetInstance().Sleep(100);
				return;
			}

            object togglePatternObject = null;

            if (base.uiElement.TryGetCurrentPattern(TogglePattern.Pattern, 
                out togglePatternObject) == true)
            {
                TogglePattern togglePattern = togglePatternObject as TogglePattern;
                Debug.Assert(togglePattern != null);
				
				togglePattern.Toggle();
			}
			else
			{
				throw new Exception("Cannot toggle checkbox");
			}
		}

        /// <summary>
        /// Gets/Sets a boolean to determine if a checkbox is checked or not,
		/// true = checked, false = unchecked, null = indeterminate
        /// </summary>
        public bool? IsChecked
        {
            get
            {
                if (this.IsAlive == false)
                {
                    throw new Exception("This UI element is not available to the user anymore.");
                }

                object togglePatternObject = null;

                if (base.uiElement.TryGetCurrentPattern(TogglePattern.Pattern,
                    out togglePatternObject) == true)
                {
                    TogglePattern togglePattern = togglePatternObject as TogglePattern;

                    if (togglePattern != null)
                    {
                        if (togglePattern.Current.ToggleState == ToggleState.Indeterminate)
                        {
                            // This checkbox is in an indeterminate state
							return null;
                        }
                        if (togglePattern.Current.ToggleState == ToggleState.Off)
                        {
                            return false;
                        }
                        else
                        {
                            // ToggleState.On
                            return true;
                        }
                    }
                }

                IntPtr hwnd = IntPtr.Zero;

                try
                {
                    hwnd = new IntPtr(this.uiElement.Current.NativeWindowHandle);
                }
                catch
                { }

                if (hwnd != IntPtr.Zero)
                {
                    StringBuilder className = new StringBuilder(256);
                    UnsafeNativeFunctions.GetClassName(hwnd, className, 256);

                    if (className.ToString() == "Button")
                    {
                        // common Win32 checkbox window
                        IntPtr result = UnsafeNativeFunctions.SendMessage(hwnd,
                            ButtonMessages.BM_GETCHECK, IntPtr.Zero, IntPtr.Zero);

                        if (result.ToInt32() == (int)ButtonMessages.BST_UNCHECKED)
                        {
                            //unchecked
                            return false;
                        }
                        else if (result.ToInt32() == (int)ButtonMessages.BST_CHECKED)
                        {
                            //checked
                            return true;
                        }
						else
						{
							return null;
						}
                    }
                }

                //return null;
				throw new Exception("Could not get the checked state of the checkbox");
            }
            set
            {
                if (value == true)
                {
                    this.Check();
                }
                else if (value == false)
                {
                    this.Uncheck();
                }
				else
				{
					this.SetIndeterminate();
				}
            }
        }
		
		/// <summary>
        /// Gets the text of the checkbox
        /// </summary>
		public string Text
		{
			get
			{
				return this.GetText();
			}
		}
		
		private AutomationPropertyChangedEventHandler UIAPropChangedEventHandler = null;
		
		/// <summary>
        /// Delegate for State Changed event
        /// </summary>
		/// <param name="sender">The checkbox that sent the event</param>
		/// <param name="newState">true if checkbox is checked, false if not checked and null if it's in an indeterminate state</param>
		public delegate void StateChanged(UIDA_CheckBox sender, bool? newState);
		internal StateChanged StateChangedHandler = null;
		
		/// <summary>
		/// Attaches/detaches a handler to state changed event
		/// </summary>
		public event StateChanged StateChangedEvent
		{
			add
			{
				try
				{
					if (this.StateChangedHandler == null)
					{
						bool? isChecked = null;
						try
						{
							isChecked = this.IsChecked;
						}
						catch {}
						if (isChecked == null)
						{
							lastValue = 2;
						}
						else
						{
							if (isChecked.Value == true)
							{
								lastValue = 1;
							}
							else
							{
								lastValue = 0;
							}
						}
						
						UIAPropChangedEventHandler = new AutomationPropertyChangedEventHandler(
							OnUIAutomationPropChangedEvent);
					
						if (base.uiElement.Current.FrameworkId == "WinForm")
						{
							Automation.AddAutomationPropertyChangedEventHandler(base.uiElement, TreeScope.Element,
									UIAPropChangedEventHandler, AutomationElement.NameProperty);
						}
						else
						{
							Automation.AddAutomationPropertyChangedEventHandler(base.uiElement, TreeScope.Element,
									UIAPropChangedEventHandler, TogglePattern.ToggleStateProperty);
						}
					}
					
					this.StateChangedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.StateChangedHandler -= value;
				
					if (this.StateChangedHandler == null)
					{
						if (this.UIAPropChangedEventHandler == null)
						{
							return;
						}
						
						System.Threading.Tasks.Task.Run(() => 
						{
							try
							{
								Automation.RemoveAutomationPropertyChangedEventHandler(base.uiElement, 
									this.UIAPropChangedEventHandler);
								UIAPropChangedEventHandler = null;
							}
							catch { }
						}).Wait(5000);
					}
				}
				catch {}
			}
		}
		
		//private DateTime lastChecked = DateTime.Now;
		private int lastValue = -1;
		
		private void OnUIAutomationPropChangedEvent(object sender, AutomationPropertyChangedEventArgs e)
		{
			if (this.StateChangedHandler == null)
			{
				return;
			}
		
			if (e.Property.Id == TogglePattern.ToggleStateProperty.Id)
			{
				int newValue = 0;
				try
				{
					newValue = (int)e.NewValue;
				}
				catch { }
				
				//DateTime currentChecked = DateTime.Now;
				//if (currentChecked - lastChecked <= TimeSpan.FromMilliseconds(100))
				//{
					if (lastValue != -1 && lastValue == newValue)
					{
						//lastChecked = currentChecked;
						return;
					}
				//}
				//lastChecked = currentChecked;
				lastValue = newValue;
				
				bool? isChecked = false;
				if (newValue == 1)
				{
					isChecked = true;
				}
				else if (newValue == 2)
				{
					isChecked = null;
				}
				
				this.StateChangedHandler(this, isChecked);
			}
			else if (e.Property.Id == AutomationElement.NameProperty.Id)
			{
				bool? isChecked = null;
				try
				{
					isChecked = this.IsChecked;
				}
				catch {}
				
				int currentValue = 0;
				if (isChecked == null)
				{
					currentValue = 2;
				}
				else if (isChecked.Value == true)
				{
					currentValue = 1;
				}
				
				if (currentValue != lastValue)
				{
					lastValue = currentValue;
					this.StateChangedHandler(this, isChecked);
				}
			}
		}
    }
}
