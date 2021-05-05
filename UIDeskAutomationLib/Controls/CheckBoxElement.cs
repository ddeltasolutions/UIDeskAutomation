using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Represents a CheckBox UI element
    /// </summary>
    public class CheckBoxElement : ElementBase
    {
        /// <summary>
        /// Creates a CheckBoxElement using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public CheckBoxElement(AutomationElement el)
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

                if ((togglePattern.Current.ToggleState == ToggleState.Indeterminate) ||
                    (togglePattern.Current.ToggleState == ToggleState.Off))
                {
                    togglePattern.Toggle();
                }

                return;
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

                    if (result.ToInt32() == (int)ButtonMessages.BST_UNCHECKED)
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

                if (togglePattern.Current.ToggleState == ToggleState.On)
                {
                    togglePattern.Toggle();
                }

                return;
            }

            // current element does not support Toggle Pattern
            // todo: use Win32 messages for common Win32 checkboxes
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

                    if (result.ToInt32() == (int)ButtonMessages.BST_CHECKED)
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
        /// Gets a boolean to determine if a checkbox is checked or not
        /// </summary>
        public bool Checked
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
                            throw new Exception("This checkbox is in an indeterminate state.");
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
                    }
                }

                return false;
            }
            set
            {
                if (value == true)
                {
                    this.Check();
                }
                else
                {
                    this.Uncheck();
                }
            }
        }
    }
}
