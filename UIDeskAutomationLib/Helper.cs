using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Automation;
using System.Drawing;
using System.Windows;

namespace UIDeskAutomationLib
{
    internal class Helper
    {
        public static string WildcardToRegex(string pattern)
        {
            return "^" + Regex.Escape(pattern)
                              .Replace(@"\*", ".*")
                              .Replace(@"\?", ".")
                       + "$";
        }

        public static IntPtr[] GetWindowsWithClass(IntPtr parentHandle, string className)
        {
            List<IntPtr> windowsList = new List<IntPtr>();
            IntPtr currentWindowHandle = IntPtr.Zero;

            do
            {
                currentWindowHandle = UnsafeNativeFunctions.FindWindowEx(parentHandle, currentWindowHandle,
                    className, null);

                if (currentWindowHandle != IntPtr.Zero)
                {
                    windowsList.Add(currentWindowHandle);
                }
            }
            while (currentWindowHandle != IntPtr.Zero);

            return windowsList.ToArray();
        }

        public static IntPtr[] GetWindows(IntPtr parentHandle, string className, 
            string windowText, bool caseSensitive)
        {
            List<IntPtr> windows = new List<IntPtr>();
            IntPtr[] windowsWithClass = Helper.GetWindowsWithClass(parentHandle, className);

            if (windowText == null)
            {
                return windowsWithClass;
            }

            if (caseSensitive == false)
            {
                windowText = windowText.ToLower();
            }

            string windowTextRegEx = Helper.WildcardToRegex(windowText);
            Regex regex = new Regex(windowTextRegEx);

            foreach (IntPtr hwnd in windowsWithClass)
            {
                StringBuilder windowTextBuilder = new StringBuilder(256);

                UnsafeNativeFunctions.GetWindowText(hwnd, windowTextBuilder, 256);
                string currentWindowText = windowTextBuilder.ToString();

                if (caseSensitive == false)
                {
                    currentWindowText = currentWindowText.ToLower();
                }

                Match match = null;

                try
                {
                    match = regex.Match(currentWindowText);
                }
                catch
                {
                    continue;
                }

                if ((match.Success == true) && (match.Value == currentWindowText))
                {
                    windows.Add(hwnd);
                }
            }

            return windows.ToArray();
        }

        //Gets a toplevel window at the specified index using wildcards in windowText
        public static Errors GetWindowAt(IntPtr parentHandle, string className, 
            string windowText, int index, bool caseSensitive, 
            out IntPtr childHandle)
        {
            if (index < 0)
            {
                childHandle = IntPtr.Zero;

                return Errors.NegativeIndex;
            }

            if (index == 0)
            {
                index = 1;
            }

            IntPtr[] windows = Helper.GetWindows(/*IntPtr.Zero*/ 
                parentHandle, className, windowText, caseSensitive);

            if ((windows == null) || (windows.Length == 0))
            {
                childHandle = IntPtr.Zero;
                return Errors.ElementNotFound;
            }

            if (index > windows.Length)
            {
                childHandle = IntPtr.Zero;
                return Errors.IndexTooBig;
            }

            if (parentHandle == IntPtr.Zero)
            {
                //sort windows list by creation time of starting process for toplevel windows
                List<WindowWithInfo> sortedWindowsList = new List<WindowWithInfo>();

                foreach (IntPtr hwnd in windows)
                {
                    WindowWithInfo windowWithInfo = new WindowWithInfo(hwnd);
                    sortedWindowsList.Add(windowWithInfo);
                }

                sortedWindowsList = sortedWindowsList.OrderBy(x => x.creationDate).ToList();
                childHandle = sortedWindowsList[index - 1].hwnd;
            }
            else
            {
                //child windows, controls
                childHandle = windows[index - 1];
            }

            return Errors.None;
        }

        public static List<AutomationElement> MatchStrings(
            AutomationElementCollection collection, string name, 
            bool bSearchByLabel, bool caseSensitive)
        {
            if (name == null)
            {
                return collection.Cast<AutomationElement>().ToList();
            }

            if (caseSensitive == false)
            {
                name = name.ToLower();
            }

            string regExName = Helper.WildcardToRegex(name);
            Regex regex = new Regex(regExName, RegexOptions.Multiline);

            List<AutomationElement> returnList = new List<AutomationElement>();

            foreach (AutomationElement el in collection)
            {
                string currentName = null;

                if (bSearchByLabel == true)
                {
                    currentName = Helper.GetElementName(el);
                }
                else
                {
                    try
                    {
                        // AutomationId ??
                        currentName = el.Current.Name;
                    }
                    catch { }
                }

                if (currentName == null)
                {
                    continue;
                }

                if (caseSensitive == false)
                {
                    currentName = currentName.ToLower();
                }

                Match match = null;
                try
                {
                    match = regex.Match(currentName);
                }
                catch
                {
                    continue;
                }

                if ((match.Success == true) && (currentName.Contains(match.Value)))
                {
                    returnList.Add(el);
                }
            }

            return returnList;
        }

        public static IntPtr GetTopLevelByProcId(int processId, 
            string className, string windowText, bool caseSensitive)
        {
            IntPtr windowHandle = IntPtr.Zero;
            uint currentProcessId = 0;

            Regex regex = null;

            if (windowText != null)
            {
                if (caseSensitive == false)
                {
                    windowText = windowText.ToLower();
                }

                string regExText = Helper.WildcardToRegex(windowText);
                regex = new Regex(regExText);
            }

            do
            {
                windowHandle = UnsafeNativeFunctions.FindWindowEx(IntPtr.Zero, windowHandle, className, null);

                UnsafeNativeFunctions.GetWindowThreadProcessId(windowHandle, out currentProcessId);

                StringBuilder windowTextBuilder = new StringBuilder(256);

                UnsafeNativeFunctions.GetWindowText(windowHandle, windowTextBuilder, 256);

                string currentWindowText = windowTextBuilder.ToString();

                if (caseSensitive == false)
                {
                    currentWindowText = currentWindowText.ToLower();
                }

                if ((int)currentProcessId == processId)
                {
                    if (windowText != null)
                    {
                        Match match = null;

                        try
                        {
                            match = regex.Match(currentWindowText);
                        }
                        catch
                        {
                            continue;
                        }

                        if ((match.Success == true) && (match.Value == currentWindowText))
                        {
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
            }
            while (windowHandle != IntPtr.Zero);

            return windowHandle;
        }

        #region Text functions
        public static string GetElementName(AutomationElement element)
        {
            if (element == null)
            {
                throw new Exception("Invalid Automation Element");
            }

            ControlType type = null;
            try
            {
                object typeObj =
                    element.GetCurrentPropertyValue(AutomationElement.ControlTypeProperty);
                type = typeObj as ControlType;
            }
            catch { }

            string name = null;

            try
            {
                object nameObject = element.GetCurrentPropertyValue(AutomationElement.NameProperty);
                name = nameObject.ToString();
            }
            catch { }
            
            //for ComboBox, Document and Edit ignore the AutomationElement.Name property
            if (type == ControlType.ComboBox || type == ControlType.Document || type == ControlType.Edit)
            {
                return null;
            }

            return name;
        }

        public static string GetElementText(AutomationElement element)
        {
            if (element == null)
            {
                throw new Exception("Invalid Automation Element");
            }

            ValuePattern valuePattern = null;
            object valuePatternObject = null;
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out valuePatternObject) == true)
            {
                valuePattern = valuePatternObject as ValuePattern;
            }

            if (valuePattern == null)
            {
                // ValuePattern not supported
                string text = null;
                try
                {
                    text = element.Current.Name;
                }
                catch { }

                return text;
            }

            // ValuePattern supported
            string value = null;
            try
            {
                value = valuePattern.Current.Value;
            }
            catch { }

            return value;
        }
        #endregion
		
		// xScreen, yScreen - screen coordinates
        /// <summary>
        /// Simulates click
        /// </summary>
        /// <param name="mouseButton">1 - Left click, 2 - Right click, 
        /// 3 - Middle click, 4 - Double click</param>
        /// <param name="xScreen"></param>
        /// <param name="yScreen"></param>
        /// <param name="keys"></param>
        /// <param name="hwnd"></param>
        public static void SimulateClickAt(int mouseButton, int xScreen, int yScreen, 
            int keys, IntPtr hwnd)
        {
            POINT ptScreen = new POINT(xScreen, yScreen);

            if (hwnd == IntPtr.Zero)
            {
                hwnd = UnsafeNativeFunctions.WindowFromPoint(ptScreen);
            }

            if (UnsafeNativeFunctions.ScreenToClient(hwnd, ref ptScreen) == false)
            {
                Engine.TraceInLogFile("Helper::SimulateClickAt - Error.");
                throw new Exception("Helper::SimulateClickAt - Error.");
            }

            // not ptScreen has client coordinates
            int coordinates = (ptScreen.Y << 16) + ptScreen.X;

            uint mouseDownMessage = 0;
            uint mouseUpMessage = 0;
            int mouseButtonConstant = 0;

            if (mouseButton == 1) // left mouse button
            {
                mouseDownMessage = WindowMessages.WM_LBUTTONDOWN;
                mouseUpMessage = WindowMessages.WM_LBUTTONUP;
                mouseButtonConstant = Win32Constants.MK_LBUTTON;
            }
            else if (mouseButton == 2) // right mouse button
            {
                mouseDownMessage = WindowMessages.WM_RBUTTONDOWN;
                mouseUpMessage = WindowMessages.WM_RBUTTONUP;
                mouseButtonConstant = Win32Constants.MK_RBUTTON;
            }
            else if (mouseButton == 3) // middle mouse button
            {
                mouseDownMessage = WindowMessages.WM_MBUTTONDOWN;
                mouseUpMessage = WindowMessages.WM_MBUTTONUP;
                mouseButtonConstant = Win32Constants.MK_MBUTTON;
            }

            if (keys == 0) // no key pressed
            {
                if (mouseButton == 4) // double click
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDOWN,
                        new IntPtr(Win32Constants.MK_LBUTTON), new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON), new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDBLCLK,
                        new IntPtr(Win32Constants.MK_LBUTTON), new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON), new IntPtr(coordinates));
                }
                else
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseDownMessage,
                        new IntPtr(mouseButtonConstant), new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseUpMessage,
                        new IntPtr(mouseButtonConstant), new IntPtr(coordinates));
                }
            }
            else if (keys == 1) // control is pressed
            {
                if (mouseButton == 4) // double click
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDOWN,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDBLCLK,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL), 
                        new IntPtr(coordinates));
                }
                else
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseDownMessage,
                        new IntPtr(mouseButtonConstant | Win32Constants.MK_CONTROL),
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseUpMessage,
                        new IntPtr(mouseButtonConstant | Win32Constants.MK_CONTROL),
                        new IntPtr(coordinates));
                }
            }
            else if (keys == 2) // shift is pressed
            {
                if (mouseButton == 4) // double click
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDOWN,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDBLCLK,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                }
                else
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseDownMessage,
                        new IntPtr(mouseButtonConstant | Win32Constants.MK_SHIFT),
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseUpMessage,
                        new IntPtr(mouseButtonConstant | Win32Constants.MK_SHIFT),
                        new IntPtr(coordinates));
                }
            }
            else if (keys == 3) // both control and shift are pressed
            {
                if (mouseButton == 4)
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDOWN,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONDBLCLK,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_LBUTTONUP,
                        new IntPtr(Win32Constants.MK_LBUTTON | Win32Constants.MK_CONTROL | Win32Constants.MK_SHIFT), 
                        new IntPtr(coordinates));
                }
                else
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseDownMessage,
                        new IntPtr(mouseButtonConstant | Win32Constants.MK_CONTROL | Win32Constants.MK_SHIFT),
                        new IntPtr(coordinates));
                    UnsafeNativeFunctions.PostMessage(hwnd, mouseUpMessage,
                        new IntPtr(mouseButtonConstant | Win32Constants.MK_CONTROL | Win32Constants.MK_SHIFT),
                        new IntPtr(coordinates));
                }
            }
            else
            {
                Engine.TraceInLogFile("Helper::SimulateClickAt - invalid argument 'keys'");
                throw new Exception("Helper::SimulateClickAt - invalid argument 'keys'");
            }
        }
		
		internal static string[] GetKeys(string text)
        {
            bool insideBrackets = false;
            List<string> keys = new List<string>();
            string currentKey = string.Empty;

            for (int i = 0; i < text.Length; i++)
            {
                if ((text[i] == '{') && (insideBrackets == false))
                {
                    insideBrackets = true;
                    currentKey += text[i];

                    continue;
                }

                if (text[i] == '}')
                {
                    if ((i < (text.Length - 1)) && (text[i + 1] == '}'))
                    {
                        currentKey += "}";
                        i++;
                    }
                    insideBrackets = false;

                    currentKey += text[i];
                    keys.Add(currentKey);
                    currentKey = string.Empty;

                    continue;
                }

                if (insideBrackets == false)
                {
                    currentKey = new string(text[i], 1);
                    keys.Add(currentKey);
                    currentKey = string.Empty;
                }
                else
                {
                    currentKey += text[i];
                }
            }

            return keys.ToArray();
        }
		
		internal static void SimulateSendKeys(string text, IntPtr hwnd)
        {
            string[] keys = Helper.GetKeys(text);

            bool bInsideBrackets = false;
            bool bShiftIsPressed = false;
            bool bCtrlIsPressed = false;
            bool bAltIsPressed = false;

            foreach (string key in keys)
            {
                if ((key == "(") && (bInsideBrackets == false))
                {
                    bInsideBrackets = true;
                    continue;
                }
                if ((key == ")") && (bInsideBrackets == true))
                {
                    // Release Control, Alt or Shift if they are pressed
                    if (bShiftIsPressed == true)
                    {
                        // Shift up
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYUP,
                            new IntPtr(VirtualKeyCodes.VK_SHIFT), IntPtr.Zero);
                        bShiftIsPressed = false;
                    }
                    if (bCtrlIsPressed == true)
                    {
                        // Ctrl up
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYUP,
                            new IntPtr(VirtualKeyCodes.VK_CONTROL), IntPtr.Zero);
                        bCtrlIsPressed = false;
                    }
                    if (bAltIsPressed == true)
                    {
                        // Alt up
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYUP,
                            new IntPtr(VirtualKeyCodes.VK_MENU), IntPtr.Zero);
                        bAltIsPressed = false;
                    }

                    bInsideBrackets = false;
                    continue;
                }

                // Press key
                if (key == "%")
                {
                    // Alt key down
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_MENU), IntPtr.Zero);
                    bAltIsPressed = true;
                }
                else if (key == "+")
                {
                    // Shift down
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_SHIFT), IntPtr.Zero);
                    bShiftIsPressed = true;
                }
                else if (key == "^")
                {
                    // Ctrl down
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_CONTROL), IntPtr.Zero);
                    bCtrlIsPressed = true;
                }
                else if ((key == "{BACKSPACE}") || (key == "{BS}") || (key == "{BKSP}"))
                {
                    // Backspace
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_BACK), IntPtr.Zero);
                }
                else if (key == "{BREAK}") // Ctrl-Break
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_CANCEL), IntPtr.Zero);
                }
                else if ((key == "{DELETE}") || (key == "{DEL}"))
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_DELETE), IntPtr.Zero);
                }
                else if (key == "{DOWN}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_DOWN), IntPtr.Zero);
                }
                else if (key == "{END}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_END), IntPtr.Zero);
                }
                else if (key == "{ENTER}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_RETURN), IntPtr.Zero);
                }
                else if (key == "{ESC}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_ESCAPE), IntPtr.Zero);
                }
                else if (key == "{HELP}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_HELP), IntPtr.Zero);
                }
                else if (key == "{HOME}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_HOME), IntPtr.Zero);
                }
                else if ((key == "{INS}") || (key == "{INSERT}"))
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_INSERT), IntPtr.Zero);
                }
                else if (key == "{LEFT}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_LEFT), IntPtr.Zero);
                }
                else if (key == "{PGDN}") // page down
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_NEXT), IntPtr.Zero);
                }
                else if (key == "{PGUP}") // page up
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_PRIOR), IntPtr.Zero);
                }
                else if (key == "{PRTSC}") // print screen
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_SNAPSHOT), IntPtr.Zero);
                }
                else if (key == "{RIGHT}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_RIGHT), IntPtr.Zero);
                }
                else if (key == "{TAB}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_TAB), IntPtr.Zero);
                }
                else if (key == "{UP}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_UP), IntPtr.Zero);
                }
                else if (key == "{F1}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F1), IntPtr.Zero);
                }
                else if (key == "{F2}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F2), IntPtr.Zero);
                }
                else if (key == "{F3}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F3), IntPtr.Zero);
                }
                else if (key == "{F4}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F4), IntPtr.Zero);
                }
                else if (key == "{F5}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F5), IntPtr.Zero);
                }
                else if (key == "{F6}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F6), IntPtr.Zero);
                }
                else if (key == "{F7}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F7), IntPtr.Zero);
                }
                else if (key == "{F8}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F8), IntPtr.Zero);
                }
                else if (key == "{F9}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F9), IntPtr.Zero);
                }
                else if (key == "{F10}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F10), IntPtr.Zero);
                }
                else if (key == "{F11}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F11), IntPtr.Zero);
                }
                else if (key == "{F12}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F12), IntPtr.Zero);
                }
                else if (key == "{F13}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F13), IntPtr.Zero);
                }
                else if (key == "{F14}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F14), IntPtr.Zero);
                }
                else if (key == "{F15}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F15), IntPtr.Zero);
                }
                else if (key == "{F16}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_F16), IntPtr.Zero);
                }
                else if (key == "{ADD}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_ADD), IntPtr.Zero);
                }
                else if (key == "{SUBTRACT}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_SUBTRACT), IntPtr.Zero);
                }
                else if (key == "{MULTIPLY}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_MULTIPLY), IntPtr.Zero);
                }
                else if (key == "{DIVIDE}")
                {
                    UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                        new IntPtr(VirtualKeyCodes.VK_DIVIDE), IntPtr.Zero);
                }
                else if (key.Length == 1) // one character string
                {
                    // Post WM_CHAR message
                    Char ch = key[0];
                    if (ch == '~')
                    {
                        // Enter
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                            new IntPtr(VirtualKeyCodes.VK_RETURN), IntPtr.Zero);
                    }
                    else
                    {
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_CHAR,
                            new IntPtr(ch), IntPtr.Zero);
                    }
                }

                // Release Alt, Shift or Control if outside brackets ()
                bool bIsAltShiftOrCtrl = (key == "%") || (key == "+") || (key == "^");
                if ((bInsideBrackets == false) && (bIsAltShiftOrCtrl == false))
                {
                    if (bShiftIsPressed == true)
                    {
                        // Shift up
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYUP,
                            new IntPtr(VirtualKeyCodes.VK_SHIFT), IntPtr.Zero);
                        bShiftIsPressed = false;
                    }
                    if (bCtrlIsPressed == true)
                    {
                        // Ctrl up
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYUP,
                            new IntPtr(VirtualKeyCodes.VK_CONTROL), IntPtr.Zero);
                        bCtrlIsPressed = false;
                    }
                    if (bAltIsPressed == true)
                    {
                        // Alt up
                        UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYUP,
                            new IntPtr(VirtualKeyCodes.VK_MENU), IntPtr.Zero);
                        bAltIsPressed = false;
                    }
                }
            }
        }
		
		private static void PressKey(IntPtr chVirtKeyCode, IntPtr hwnd)
        {
            UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYDOWN,
                chVirtKeyCode, IntPtr.Zero);

            UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_KEYUP,
                chVirtKeyCode, IntPtr.Zero);
        }

        internal static IntPtr GetTopLevelByProcName(string processName, 
            string className, string windowText, int index, bool caseSensitive)
        {
            processName = processName.ToLower();

            if (caseSensitive == false)
            {
                windowText = windowText.ToLower();
            }

            IntPtr windowHandle = IntPtr.Zero;
            uint currentProcessId = 0;

            Regex regex = null;

            if (windowText != null)
            {
                string regExText = Helper.WildcardToRegex(windowText);
                regex = new Regex(regExText);
            }

            do
            {
                windowHandle = UnsafeNativeFunctions.FindWindowEx(IntPtr.Zero, windowHandle, className, null);

                if (windowHandle == IntPtr.Zero)
                {
                    break;
                }

                if (UnsafeNativeFunctions.IsWindowVisible(windowHandle) == false)
                { 
                    // continue to next sibling window if current window is not visible
                    continue;
                }

                UnsafeNativeFunctions.GetWindowThreadProcessId(windowHandle, out currentProcessId);

                Process currentProcess = Process.GetProcessById((int)currentProcessId);

                string currentProcessName = currentProcess.ProcessName.ToLower() + ".exe";

                StringBuilder windowTextBuilder = new StringBuilder(256);

                UnsafeNativeFunctions.GetWindowText(windowHandle, windowTextBuilder, 256);

                string currentWindowText = windowTextBuilder.ToString();

                if (caseSensitive == false)
                {
                    currentWindowText = currentWindowText.ToLower();
                }

                if (currentProcessName == processName)
                {
                    if (windowText != null)
                    {
                        Match match = null;

                        try
                        {
                            match = regex.Match(currentWindowText);
                        }
                        catch 
                        {
                            continue;
                        }

                        if ((match.Success == true) && (match.Value == currentWindowText))
                        {
                            index--;
                        }
                    }
                    else
                    {
                        index--;
                    }

                    if (index == 0)
                    {
                        break;
                    }
                }
            }
            while (windowHandle != IntPtr.Zero);

            if (index == 0)
            {
                return windowHandle;
            }
            else
            {
                // not found
                return IntPtr.Zero;
            }
        }

        //Returns true is elements are identical, false otherwise.
        internal static bool CompareAutomationElements(AutomationElement el1,
            AutomationElement el2)
        {
            int[] runtimeId1 = null;
            int[] runtimeId2 = null;

            try
            {
                runtimeId1 = el1.GetRuntimeId();
                runtimeId2 = el2.GetRuntimeId();
            }
            catch { }

            if ((runtimeId1 == null) || (runtimeId2 == null))
            {
                return (el1 == el2);
            }

            if (runtimeId1.Length != runtimeId2.Length)
            {
                return false;
            }

            bool different = false;

            for (int i = 0; i < runtimeId1.Length; i++)
            {
                int id1 = runtimeId1[i];
                int id2 = runtimeId2[i];

                if (id1 != id2)
                {
                    different = true;
                    break;
                }
            }

            return !different;
        }

        internal static ValuePattern GetValuePattern(AutomationElement uiElement)
        {
            object valuePatternObj = null;

            uiElement.TryGetCurrentPattern(ValuePattern.Pattern,
                out valuePatternObj);

            ValuePattern valuePattern = valuePatternObj as ValuePattern;

            return valuePattern;
        }

        internal static RangeValuePattern GetRangeValuePattern(
            AutomationElement uiElement)
        {
            object rangeValuePatternObj = null;

            uiElement.TryGetCurrentPattern(RangeValuePattern.Pattern,
                out rangeValuePatternObj);

            RangeValuePattern rangeValuePattern = 
                rangeValuePatternObj as RangeValuePattern;

            return rangeValuePattern;
        }
        
        private static Bitmap CreateWindowSnapshot(IntPtr hWnd, AutomationElement element)
        {
            RECT rect;
            if (UnsafeNativeFunctions.GetWindowRect(hWnd, out rect) == false)
            {
                return null;
            }
            
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;
            int left = 0;
            int top = 0;

            IntPtr hdc = UnsafeNativeFunctions.GetWindowDC(hWnd);
            if (hdc == IntPtr.Zero)
            {
                return null;
            }
            IntPtr hdcMem = UnsafeNativeFunctions.CreateCompatibleDC(hdc);
            if (hdcMem == IntPtr.Zero)
            {
                return null;
            }
            IntPtr hbmp = UnsafeNativeFunctions.CreateCompatibleBitmap(hdc, width, height);
            if (hbmp == IntPtr.Zero)
            {
                return null;
            }
            if (UnsafeNativeFunctions.SelectObject(hdcMem, hbmp) == IntPtr.Zero)
            {
                return null;
            }
            if (UnsafeNativeFunctions.BitBlt(hdcMem, 0, 0, width, height, hdc, left, top, 
                TernaryRasterOperations.SRCCOPY | TernaryRasterOperations.CAPTUREBLT) == false)
            {
                return null;
            }

            Bitmap bitmap = System.Drawing.Image.FromHbitmap(hbmp);

            UnsafeNativeFunctions.DeleteObject(hbmp);
            UnsafeNativeFunctions.DeleteDC(hdcMem);
            UnsafeNativeFunctions.ReleaseDC(hWnd, hdc);
            return bitmap;
        }
        
        private static Bitmap CreateSnapshot(IntPtr hWnd, int left, int top, int right, int bottom)
        {
            POINT pt;
            pt.X = left;
            pt.Y = top;
            if (hWnd != IntPtr.Zero && UnsafeNativeFunctions.ScreenToClient(hWnd, ref pt) == false)
            {
                return null;
            }
            if (pt.X < 0 || pt.Y < 0)
            {
                Engine.TraceInLogFile("Out of client area");
                return null;
            }
            int width = right - left;
            int height = bottom - top;

            //HDC hdc = GetWindowDC(hWnd);
            IntPtr hdc = UnsafeNativeFunctions.GetDC(hWnd);
            if (hdc == IntPtr.Zero)
            {
                return null;
            }
            IntPtr hdcMem = UnsafeNativeFunctions.CreateCompatibleDC(hdc);
            if (hdcMem == IntPtr.Zero)
            {
                return null;
            }
            IntPtr hbmp = UnsafeNativeFunctions.CreateCompatibleBitmap(hdc, width, height);
            if (hbmp == IntPtr.Zero)
            {
                return null;
            }
            if (UnsafeNativeFunctions.SelectObject(hdcMem, hbmp) == IntPtr.Zero)
            {
                return null;
            }
            if (UnsafeNativeFunctions.BitBlt(hdcMem, 0, 0, width, height, hdc, pt.X, pt.Y, 
                TernaryRasterOperations.SRCCOPY | TernaryRasterOperations.CAPTUREBLT) == false)
            {
                return null;
            }

            Bitmap bitmap = System.Drawing.Image.FromHbitmap(hbmp);

            UnsafeNativeFunctions.DeleteObject(hbmp);
            UnsafeNativeFunctions.DeleteDC(hdcMem);
            UnsafeNativeFunctions.ReleaseDC(hWnd, hdc);
            return bitmap;
        }
        
        private static AutomationElement GetParentWindow(AutomationElement element)
        {
            TreeWalker tw = TreeWalker.ControlViewWalker;
            AutomationElement root = AutomationElement.RootElement;
            AutomationElement crtElement = element;
            
            if (crtElement == null || CompareAutomationElements(crtElement, root) == true)
            {
                return null;
            }
            
            do
            {
                crtElement = tw.GetParent(crtElement);
                if (crtElement == null || CompareAutomationElements(crtElement, root) == true)
                {
                    return null;
                }
                
                IntPtr hWnd = new IntPtr(crtElement.Current.NativeWindowHandle);
                if (hWnd != IntPtr.Zero)
                {
                    return crtElement;
                }
            }
            while (true);
        }
        
        internal static Bitmap CaptureElement(AutomationElement element)
        {
            if (element.Current.ControlType == ControlType.Menu || 
                element.Current.ControlType == ControlType.MenuItem)
            {
                return null;
            }
            
            if (element.Current.ControlType == ControlType.Spinner && 
                element.Current.FrameworkId == "WinForm")
            {
                element = TreeWalker.ControlViewWalker.GetParent(element);
            }
			
			Bitmap bitmap = null;
			IntPtr hWnd = new IntPtr(element.Current.NativeWindowHandle);
			if (hWnd != IntPtr.Zero)
			{
				bitmap = CreateWindowSnapshot(hWnd, element);
				if (bitmap != null)
				{
					return bitmap;
				}
			}
			AutomationElement crtParent = element;
			if (hWnd == IntPtr.Zero)
			{
				crtParent = GetParentWindow(element);
				if (crtParent == null)
				{
					return null;
				}
				hWnd = new IntPtr(crtParent.Current.NativeWindowHandle);
			}
			
			System.Windows.Rect rect = element.Current.BoundingRectangle;
			
			while (bitmap == null)
			{
				bitmap = CreateSnapshot(hWnd, (int)rect.Left, (int)rect.Top, 
					(int)rect.Right, (int)rect.Bottom);
				if (bitmap == null)
				{
					crtParent = GetParentWindow(crtParent);
					if (crtParent == null)
					{
						return null;
					}
					hWnd = new IntPtr(crtParent.Current.NativeWindowHandle);
				}
			}
			
			return bitmap;
        }
		
		internal static Bitmap CaptureVisibleElement(AutomationElement element, Rect? cropRect = null)
		{
			Rect rect = element.Current.BoundingRectangle;
			System.Drawing.Size size = new System.Drawing.Size((int)rect.Width, (int)rect.Height);
			if (cropRect.HasValue)
			{
				size = new System.Drawing.Size((int)cropRect.Value.Width, (int)cropRect.Value.Height);
			}
			Bitmap bitmap = new Bitmap(size.Width, size.Height);
			
			using (Graphics g = Graphics.FromImage(bitmap))
			{
				if (cropRect.HasValue)
				{
					g.CopyFromScreen((int)(rect.Left + cropRect.Value.Left), (int)(rect.Top + cropRect.Value.Top), 0, 0, size);
				}
				else
				{
					g.CopyFromScreen((int)rect.Left, (int)rect.Top, 0, 0, size);
				}
			}
			
			return bitmap;
		}
		
		internal static Bitmap CropImage(Bitmap image, Rectangle cropRect)
		{
			Bitmap result = new Bitmap(cropRect.Width, cropRect.Height);

			using(Graphics g = Graphics.FromImage(result))
			{
				g.DrawImage(image, new Rectangle(0, 0, result.Width, result.Height), 
						cropRect, GraphicsUnit.Pixel);
			}
			
			return result;
		}
    }

    internal class WindowWithInfo
    {
        public IntPtr hwnd = IntPtr.Zero;

        public DateTime creationDate = DateTime.MinValue;

        public WindowWithInfo(IntPtr windowHandle)
        {
            this.hwnd = windowHandle;
            uint processId = 0;

            UnsafeNativeFunctions.GetWindowThreadProcessId(this.hwnd, out processId);

            Process process = Process.GetProcessById(Convert.ToInt32(processId));
            this.creationDate = process.StartTime;
        }
    }
}
