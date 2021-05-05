﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Xml;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// dDeltaSolutions.PSLib namespace contains classes for automating 
    /// Windows user interface elements.
    /// </summary>
    [System.Runtime.CompilerServices.CompilerGenerated]
    class NamespaceDoc
    { }

    /// <summary>
    /// Main entry point class. You need only one instance of this class.
    /// </summary>
    public class Engine
    {
        private static int wait = 5000; //default wait 5 seconds
        private static string logFileName = string.Empty;
        private static bool throwExceptionsWhenSearch = true;

        private static Engine instance = null;

        /// <summary>
        /// Constructor
        /// </summary>
        public Engine()
        {
            //EvaluationFrm evaluationForm = new EvaluationFrm();
            //evaluationForm.ShowDialog();

            Engine.instance = this;
            return;
        }

        /// <summary>
        /// Gets/Sets Timeout period (in milliseconds)
        /// </summary>
        public int Timeout
        {
            get
            {
                return Engine.wait;
            }
            set
            {
                Engine.wait = value;
            }
        }

        /// <summary>
        /// Gets/Sets Log File path
        /// </summary>
        public string LogFile
        { 
            get
            {
                return Engine.logFileName;
            }
            set
            {
                Engine.logFileName = value;

                try
                {
                    if (File.Exists(Engine.logFileName))
                    {
                        //clear log file content
                        File.WriteAllText(Engine.logFileName, string.Empty);
                    }
                }
                catch { }
            }
        }
        
        internal static bool ThrowExceptionsWhenSearch
        {
            get
            {
                return Engine.throwExceptionsWhenSearch;
            }
            set
            {
                Engine.throwExceptionsWhenSearch = value;
            }
        }

        /// <summary>
        /// Gets/Sets whether script should raise exceptions when elements are not found.
        /// </summary>
        public bool ThrowExceptionsForSearchFunctions
        {
            get
            {
                return Engine.throwExceptionsWhenSearch;
            }
            set
            {
                Engine.throwExceptionsWhenSearch = value;
            }
        }

        /// <summary>
        /// Writes a message in the log file
        /// </summary>
        /// <param name="sMessage">message</param>
        internal static void TraceInLogFile(string sMessage)
        {
            if (Engine.logFileName.Length == 0)
            {
                return;
            }

            try
            {
                string sLineToWrite = DateTime.Now.ToString("G") + ": " + sMessage + Environment.NewLine;
                File.AppendAllText(Engine.logFileName, sLineToWrite);
            }
            catch { }
        }

        /// <summary>
        /// Writes a message in the log file
        /// </summary>
        /// <param name="sMessage">message to write in log file</param>
        public void WriteInLogFile(string sMessage)
        {
            Engine.TraceInLogFile(sMessage);

            return;
        }

        /// <summary>
        /// Gets a top-level window element given its class name and/or window text
        /// </summary>
        /// <param name="className">Window class name (optional)</param>
        /// <param name="windowText">Window text (optional), wildcards can be used</param>
        /// <param name="index">index (optional)</param>
        /// <param name="caseSensitive">searches the window text case sensitive, default true</param>
        /// <returns>Window</returns>
        public Window GetTopLevel(string className, string windowText, int index = 0,
            bool caseSensitive = true)
        {
            if (index < 0)
            {
                Engine.TraceInLogFile("GetTopLevel method - index must be positive number");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GetTopLevel method - index must be positive number");
                }
                else
                {
                    return null;
                }
            }

            if (index == 0)
            {
                index = 1;
            }

            int nWaitMs = Engine.wait;
            IntPtr hwnd = IntPtr.Zero;

            Errors error = Errors.None;

            while (nWaitMs > 0)
            {
                error = Helper.GetWindowAt(IntPtr.Zero, className, windowText, 
                    index, caseSensitive, out hwnd);

                if ((hwnd != IntPtr.Zero) && (error == Errors.None))
                {
                    break; //found!
                }

                nWaitMs -= 100; //wait 100 milliseconds

                Thread.Sleep(100);
            }

            if (hwnd == IntPtr.Zero)
            {
                if (error == Errors.IndexTooBig)
                {
                    Engine.TraceInLogFile("GetTopLevel method - Index too big");

                    if (Engine.ThrowExceptionsWhenSearch == true)
                    {
                        throw new Exception("GetTopLevel method - Index too big");
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    Engine.TraceInLogFile("GetTopLevel method - Window not found");

                    if (Engine.ThrowExceptionsWhenSearch == true)
                    {
                        throw new Exception("GetTopLevel method - Window not found");
                    }
                    else
                    {
                        return null;
                    }
                }
            }

            return new Window(hwnd);
        }

        /// <summary>
        /// Gets the top level window of the specified process with the specified 
        /// class name, window text and index
        /// </summary>
        /// <param name="processName">process name (ex: "myprocess.exe")</param>
        /// <param name="className">Window class name</param>
        /// <param name="windowText">Window text, wildcards can be used</param>
        /// <param name="index">index</param>
        /// <param name="caseSensitive">searches the window text case sensitive, default true</param>
        /// <returns>Top level Window</returns>
        public Window GetTopLevelByProcName(string processName, string className = null,
            string windowText = null, int index = 0, bool caseSensitive = true)
        {
            int nWaitMs = Engine.wait;
            IntPtr windowHandle = IntPtr.Zero;

            if (index < 0)
            { 
                Engine.TraceInLogFile("GetTopLevelByProcName failed: negative index");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GetTopLevelByProcName failed - negative index");
                }
                else
                {
                    return null;
                }
            }

            if (index == 0)
            {
                index = 1;
            }

            while (nWaitMs > 0)
            {
                windowHandle = Helper.GetTopLevelByProcName(processName, 
                    className, windowText, index, caseSensitive);

                if (windowHandle != IntPtr.Zero)
                {
                    break; //found!
                }

                nWaitMs -= 100; //wait 100 milliseconds

                Thread.Sleep(100);
            }

            if (windowHandle == IntPtr.Zero)
            {
                Engine.TraceInLogFile("GetTopLevelByProcName - window not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GetTopLevelByProcName method - window not found");
                }
                else
                {
                    return null;
                }
            }

            return new Window(windowHandle);
        }

        /// <summary>
        /// Get the toplevel windows with the specified class name and/or window text.
        /// </summary>
        /// <param name="className">Window class name</param>
        /// <param name="windowText">Window text, wildcards can be used</param>
        /// <param name="caseSensitive">searches the windows text case sensitive, default true</param>
        /// <returns>Windows collection</returns>
        public Window[] GetTopLevelWindows(string className = null,
            string windowText = null, bool caseSensitive = true)
        {
            IntPtr[] handles = Helper.GetWindows(IntPtr.Zero, className, 
                windowText, caseSensitive);

            List<Window> windows = new List<Window>();

            foreach (IntPtr handle in handles)
            {
                Window wnd = new Window(handle);
                windows.Add(wnd);
            }

            return windows.ToArray();
        }
        

        /// <summary>
        /// Get the toplevel window of the specified process with the specified class name and/or window text.
        /// </summary>
        /// <param name="processId">Process Id</param>
        /// <param name="className">Window class name</param>
        /// <param name="windowText">Window text, wildcards can be used</param>
        /// <param name="caseSensitive">searches the window text case sensitive, default true</param>
        /// <returns>Window</returns>
        public Window GetTopLevelByProcId(int processId, string className = null, 
            string windowText = null, bool caseSensitive = true)
        {
            Process process = null;
            try
            {
                process = Process.GetProcessById(processId);
            }
            catch { }

            if (process == null)
            {
                Engine.TraceInLogFile("GetTopLevelByProcId method - process with this id doesn't exist");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GetTopLevelByProcId method - process with this id doesn't exist");
                }
                else
                {
                    return null;
                }
            }

            int nWaitMs = Engine.wait;
            IntPtr windowHandle = IntPtr.Zero;

            while (nWaitMs > 0)
            {
                windowHandle = Helper.GetTopLevelByProcId(processId, className, 
                    windowText, caseSensitive);

                if (windowHandle != IntPtr.Zero)
                {
                    break; //found!
                }

                nWaitMs -= 100; //wait 100 milliseconds

                Thread.Sleep(100);
            }

            if (windowHandle == IntPtr.Zero)
            {
                Engine.TraceInLogFile("GetTopLevelByProcId method - window not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GetTopLevelByProcId method - window not found");
                }
                else
                {
                    return null;
                }
            }

            return new Window(windowHandle);
        }

        private IntPtr FindWindow(string sClassName, string sText, int nIndex = 0)
        {
            IntPtr hwnd = IntPtr.Zero;

            if (nIndex <= 1)
            {
                hwnd = UnsafeNativeFunctions.FindWindowEx(IntPtr.Zero, IntPtr.Zero, sClassName, sText);
            }
            else
            {
                IntPtr hwndTemp = IntPtr.Zero;

                while (nIndex >= 1)
                {
                    hwnd = UnsafeNativeFunctions.FindWindowEx(IntPtr.Zero, hwndTemp, sClassName, sText);

                    if (hwnd == IntPtr.Zero)
                    {
                        return IntPtr.Zero;
                    }

                    hwndTemp = hwnd;
                    nIndex--;
                }
            }

            return hwnd;
        }

        /// <summary>
        /// Gets a Pane ui element representing the main desktop.
        /// </summary>
        /// <returns>a Pane element that represents the desktop</returns>
        public Pane GetDesktopPane()
        {
            AutomationElement desktopPane = AutomationElement.RootElement;

            return new Pane(desktopPane);
        }

        /// <summary>
        /// Sends keystrokes to the window which currently has the focus.
		/// This function behaves like the .NET 'SendKeys.Send(string)' function.
		/// More details can be found here: https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys.send?view=netframework-4.8
        /// </summary>
        /// <param name="text">keys to send</param>
        public void SendKeys(string text)
        {
            System.Windows.Forms.SendKeys.SendWait(text);
        }
		
		/// <summary>
        /// Simulates keystrokes. 
		/// This function behaves like the .NET 'SendKeys.Send(string)' function.
        /// </summary>
        /// <param name="text">keys to send</param>
        public void SimulateSendKeys(string text)
        {
            // Get focused window from foreground window ////
            IntPtr hwndForeground = UnsafeNativeFunctions.GetForegroundWindow();
            if (hwndForeground == IntPtr.Zero)
            {
                Engine.TraceInLogFile("SimulateSendKeys failed: Cannot get foreground window.");
                throw new Exception("SimulateSendKeys failed: Cannot get foreground window.");
            }

            uint processId = 0;
            uint foregroundThreadId = 
                UnsafeNativeFunctions.GetWindowThreadProcessId(hwndForeground, out processId);

            uint currentThreadId = UnsafeNativeFunctions.GetCurrentThreadId();

            bool bResult = UnsafeNativeFunctions.AttachThreadInput(
                currentThreadId, foregroundThreadId, true);
            Debug.Assert(bResult == true);

            IntPtr hwnd = UnsafeNativeFunctions.GetFocus();

            bResult = UnsafeNativeFunctions.AttachThreadInput(
                currentThreadId, foregroundThreadId, false);
            Debug.Assert(bResult == true);
            ///////

            if (hwnd == IntPtr.Zero)
            {
                hwnd = hwndForeground;
            }

            if (hwnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("SimulateSendKeys - Cannot get the active window");
                throw new Exception("SimulateSendKeys - Cannot get the active window");
            }

            Helper.SimulateSendKeys(text, hwnd);
        }

        // x, y - screen coordinates  
        // mouse button, 1 - left button, 2 - right button, 3 - middle button
        //public void ClickScreenCoordinatesAt(int x, int y, int mouseButton = 1)
        /// <summary>
        /// Left mouse click at screen coordinates
        /// </summary>
        /// <param name="x">x - screen coordinate</param>
        /// <param name="y">y - screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift pressed</param>
        public void ClickScreenCoordinatesAt(int x, int y, int keys = 0)
        {
            // left click
            SendInputClass.ClickLeftMouseButton(x, y, keys);
        }

        // x, y - screen coordinates  
        /// <summary>
        /// Right mouse button click at specified screen coordinates
        /// </summary>
        /// <param name="x">x - screen coordinate</param>
        /// <param name="y">y - screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift pressed</param>
        public void RightClickScreenCoordinatesAt(int x, int y, int keys = 0)
        {
            // right click
            SendInputClass.ClickRightMouseButton(x, y, keys);
        }

        // x, y - screen coordinates  
        /// <summary>
        /// Middle mouse button click at specified screen coordinates
        /// </summary>
        /// <param name="x">x - screen coordinate</param>
        /// <param name="y">y - screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift pressed</param>
        public void MiddleClickScreenCoordinatesAt(int x, int y, int keys = 0)
        {
            // middle button click
            SendInputClass.ClickMiddleMouseButton(x, y, keys);
        }

        // x, y - screen coordinates 
        /// <summary>
        /// Left mouse button double click at specified screen coordinates
        /// </summary>
        /// <param name="x">x - screen coordinate</param>
        /// <param name="y">y - screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift pressed</param>
        public void DoubleClickAt(int x, int y, int keys = 0)
        {
            SendInputClass.DoubleClick(x, y, keys);
        }
		
		// simulate click functions

        /// <summary>
        /// Simulates left mouse button click
        /// </summary>
        /// <param name="x">x screen coordinate</param>
        /// <param name="y">y screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        public void SimulateClickAt(int x, int y, int keys = 0)
        {
            Helper.SimulateClickAt(1, x, y, keys, IntPtr.Zero);
        }

        /// <summary>
        /// Simulates right mouse button click
        /// </summary>
        /// <param name="x">x screen coordinate</param>
        /// <param name="y">y screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        public void SimulateRightClickAt(int x, int y, int keys = 0)
        {
            Helper.SimulateClickAt(2, x, y, keys, IntPtr.Zero);
        }

        /// <summary>
        /// Simulates middle mouse button click
        /// </summary>
        /// <param name="x">x screen coordinate</param>
        /// <param name="y">y screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        public void SimulateMiddleClickAt(int x, int y, int keys = 0)
        {
            Helper.SimulateClickAt(3, x, y, keys, IntPtr.Zero);
        }

        /// <summary>
        /// Simulates left mouse button double click
        /// </summary>
        /// <param name="x">x screen coordinate</param>
        /// <param name="y">y screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        public void SimulateDoubleClickAt(int x, int y, int keys = 0)
        {
            Helper.SimulateClickAt(4, x, y, keys, IntPtr.Zero);
        }
		
		// x, y - screen coordinates 
        /// <summary>
        /// Moves mouse pointer at the specified screen coordinates
        /// </summary>
        /// <param name="x">x - screen coordinate</param>
        /// <param name="y">y - screen coordinate</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift pressed</param>
        public void MoveMouse(int x, int y, int keys = 0)
        {
            SendInputClass.MoveMousePointer(x, y, keys);
        }
		
        /// <summary>
        /// Scrolls mouse wheel up with the specified number of wheel ticks
        /// </summary>
        /// <param name="wheelTicks">number of wheel ticks</param>
        public void MouseScrollUp(uint wheelTicks)
        {
            SendInputClass.MouseScroll(wheelTicks);
        }
		
		/// <summary>
        /// Scrolls mouse wheel down with the specified number of wheel ticks
        /// </summary>
        /// <param name="wheelTicks">number of wheel ticks</param>
        public void MouseScrollDown(uint wheelTicks)
        {
            SendInputClass.MouseScroll((uint)((-1) * wheelTicks));
        }

        /// <summary>
        /// Gets the currently Engine instance. Should not be used by user.
        /// It is for internal purposes.
        /// </summary>
        /// <returns></returns>
        internal static Engine GetInstance()
        {
            return Engine.instance;
        }

        /// <summary>
        /// Hangs execution untill a certain property of an element reaches the specified value.
		/// For example, if you want to wait until a window closes: engine.WaitUntil(window, "IsAlive", "==", false);
        /// </summary>
        /// <param name="element">Element</param>
        /// <param name="property">Property name. Can be: IsAlive, Text and Value</param>
        /// <param name="comparison">Comparison. Can be ==, !=, &lt; (less than), &gt; (greater than), &lt;= (less or equal), &gt;= (greater or equal)</param>
        /// <param name="value">Value. Depends on the property name. Can be boolean (for IsAlive), a text (for Text) and a real number (for Value)</param>
        /// <param name="timeOut">Time out for waiting. Default 5 minutes.</param>
        /// <returns>true if timed out, false otherwise</returns>
        public bool WaitUntil(ElementBase element, string property, 
            string comparison, object value, int timeOut = 300000)
        {
            bool timedOut = false;
            int timeoutTemp = 0;

            if (property == "IsAlive")
            {
                #region IsAlive property
                bool valueBool = false;

                try
                {
                    valueBool = Convert.ToBoolean(value);
                }
                catch 
                {
                    Engine.TraceInLogFile("WaitUntil - value should be a boolean");
                    throw new Exception("WaitUntil - value should be a boolean");
                }

                if (comparison == "==")
                {
                    do
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");

                            timedOut = true;
                            break;
                        }
                    }
                    while (element.IsAlive != valueBool);
                }
                else if (comparison == "!=")
                {
                    do
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");

                            timedOut = true;
                            break;
                        }
                    }
                    while (element.IsAlive == valueBool);
                }
                #endregion
            }
            else if (property == "Text")
            {
                #region Text property
                string valueText = value as string;

                if (valueText == null)
                {
                    Engine.TraceInLogFile("WaitUntil - value must be a string");
                    throw new Exception("WaitUntil - value must be a string");
                }

                if (comparison == "==")
                {
                    do
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");

                            timedOut = true;
                            break;
                        }
                    }
                    while (element.GetText() != valueText);
                }
                else if (comparison == "!=")
                {
                    do
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");

                            timedOut = true;
                            break;
                        }
                    }
                    while (element.GetText() == valueText);
                }
                #endregion
            }
            else if (property == "Value")
            {
                #region Value property
                RangeValuePattern rangeValuePattern =
                    Helper.GetRangeValuePattern(element.uiElement);

                if (rangeValuePattern == null)
                {
                    Engine.TraceInLogFile("WaitUntil - Element does not support RangeValuePattern");
                    throw new Exception("WaitUntil - Element does not support RangeValuePattern");
                }

                double valueDouble = 0.0;

                try
                {
                    valueDouble = Convert.ToDouble(value);
                }
                catch
                {
                    Engine.TraceInLogFile("WaitUntil - value should be a number");
                    throw new Exception("WaitUntil - value should be a number");
                }

                if (comparison == "==")
                {
                    while (rangeValuePattern.Current.Value != valueDouble)
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");

                            timedOut = true;
                            break;
                        }
                    }
                }
                else if (comparison == "!=")
                {
                    while (rangeValuePattern.Current.Value == valueDouble)
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");

                            timedOut = true;
                            break;
                        }
                    }
                }
                else if (comparison == "<")
                {
                    while (rangeValuePattern.Current.Value >= valueDouble)
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");

                            timedOut = true;
                            break;
                        }
                    }
                }
                else if (comparison == "<=")
                {
                    while (rangeValuePattern.Current.Value > valueDouble)
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");

                            timedOut = true;
                            break;
                        }
                    }
                }
                else if (comparison == ">")
                {
                    while (rangeValuePattern.Current.Value <= valueDouble)
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");

                            timedOut = true;
                            break;
                        }
                    }
                }
                else if (comparison == ">=")
                {
                    while (rangeValuePattern.Current.Value < valueDouble)
                    {
                        Thread.Sleep(100);
                        timeoutTemp += 100;

                        if (timeoutTemp >= timeOut)
                        {
                            Engine.TraceInLogFile("WaitUntil - timeout reached");

                            timedOut = true;
                            break;
                        }
                    }
                }
                else
                {
                    Engine.TraceInLogFile("WaitUntil - comparison sign not recognized");
                    throw new Exception("WaitUntil - comparison sign not recognized");
                }
                #endregion
            }
            else
            {
                Engine.TraceInLogFile("WaitUntil - property not supported");
                throw new Exception("WaitUntil - property not supported");
            }

            return timedOut;
        }

        /// <summary>
        /// Starts a new process
        /// </summary>
		/// <param name="processName">name of the executable file</param>
        /// <returns>Id of the new started process</returns>
		public int StartProcess(string processName)
		{
			try
			{
				Process proc = Process.Start(processName);
				return proc.Id;
			}
			catch (Exception ex)
			{
				Engine.TraceInLogFile("Cannot start process: " + ex.Message);
				throw new Exception("Cannot start process: " + ex.Message);
			}
		}
		
		/// <summary>
        /// Starts a new process and waits for input idle
        /// </summary>
		/// <param name="processName">name of the executable file</param>
        /// <returns>Id of the new started process</returns>
		public int StartProcessAndWaitForInputIdle(string processName)
		{
			try
			{
				Process proc = Process.Start(processName);
				proc.WaitForInputIdle();
				return proc.Id;
			}
			catch (Exception ex)
			{
				Engine.TraceInLogFile("Cannot start process: " + ex.Message);
				throw new Exception("Cannot start process: " + ex.Message);
			}
		}
		
		/// <summary>
        /// Starts a new process with arguments
        /// </summary>
		/// <param name="processName">name of the executable file</param>
		/// <param name="arguments">command line arguments</param>
        /// <returns>Id of the new started process</returns>
		public int StartProcess(string processName, string arguments)
		{
			try
			{
				Process proc = Process.Start(processName, arguments);
				return proc.Id;
			}
			catch (Exception ex)
			{
				Engine.TraceInLogFile("Cannot start process: " + ex.Message);
				throw new Exception("Cannot start process: " + ex.Message);
			}
		}
		
		/// <summary>
        /// Starts a new process with arguments and waits for input idle
        /// </summary>
		/// <param name="processName">name of the executable file</param>
		/// <param name="arguments">command line arguments</param>
        /// <returns>Id of the new started process</returns>
		public int StartProcessAndWaitForInputIdle(string processName, string arguments)
		{
			try
			{
				Process proc = Process.Start(processName, arguments);
				proc.WaitForInputIdle();
				return proc.Id;
			}
			catch (Exception ex)
			{
				Engine.TraceInLogFile("Cannot start process: " + ex.Message);
				throw new Exception("Cannot start process: " + ex.Message);
			}
		}
		
		/// <summary>
        /// Blocks the calling thread for a specified period of time
        /// </summary>
		/// <param name="milliseconds">number of milliseconds</param>
        public void Sleep(int milliseconds)
		{
			try
			{
				Thread.Sleep(milliseconds);
			}
			catch (Exception ex)
			{
				Engine.TraceInLogFile("Sleep failed: " + ex.Message);
				throw new Exception("Sleep failed: " + ex.Message);
			}
		}
    }

    internal enum Errors
    {
        None,
        ElementNotFound,
        IndexTooBig,
        NegativeIndex
    }

    internal enum MouseButtons
    {
        LeftButton = 1, RightButton, Wheel
    }
}
