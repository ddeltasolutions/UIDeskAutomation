using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using System.Globalization;
using System.Runtime.InteropServices;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Calendar UI element. It works only for WPF and Win32 calendar.
    /// </summary>
    public class UIDA_Calendar : ElementBase
    {
        /// <summary>
        /// Creates a Calendar using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_Calendar(AutomationElement el)
        {
            base.uiElement = el;
        }
        
        /// <summary>
        /// Gets the selected dates in calendar. For WPF calendar it gets all selected dates.
        /// For Win32 calendar, if a range of dates is selected, it gets the first and the last
        /// date of the range. If one date is selected it gets the selected date.
        /// </summary>
        /// <returns>Array of selected date/dates</returns>
        public DateTime[] SelectedDates
        {
            get
            {
                if (uiElement.Current.FrameworkId == "WPF")
                {
                    List<DateTime> result = new List<DateTime>();
                    
                    object selectionPatternObj = null;
                    uiElement.TryGetCurrentPattern(SelectionPattern.Pattern, out selectionPatternObj);
                    SelectionPattern selectionPattern = selectionPatternObj as SelectionPattern;
                    if (selectionPattern != null)
                    {
                        AutomationElement[] selection = selectionPattern.Current.GetSelection();
                        foreach (AutomationElement selectedElement in selection)
                        {
                            try
                            {
                                string name = selectedElement.Current.Name;
                                DateTime date = DateTime.Parse(name, CultureInfo.CurrentCulture);
                                result.Add(date);
                            }
                            catch { }
                        }
                    }
                    
                    return result.ToArray();
                }
                else if (uiElement.Current.FrameworkId == "Win32")
                {
                    return GetSelection(new IntPtr(uiElement.Current.NativeWindowHandle));
                }
                
                throw new Exception("Not supported");
            }
        }
        
        /// <summary>
        /// Deselects other selected dates and selects the specified date.
        /// </summary>
        /// <param name="date">date to select</param>
        public void SelectDate(DateTime date)
        {
            if (uiElement.Current.FrameworkId == "WPF")
            {
                SetSelectedDate(date, false);
            }
            else if (uiElement.Current.FrameworkId == "Win32")
            {
                SetSelectedDate(new IntPtr(uiElement.Current.NativeWindowHandle), date);
            }
            else
            {
                throw new Exception("Not supported");
            }
        }
        
        /// <summary>
        /// Deselects other selected dates and selects the specified range.
        /// For WPF calendar all dates to be selected should be specified in parameter array.
        /// For Win32 calendar only the first and the last date of the range should be specified in the parameter array.
        /// </summary>
        /// <param name="dates">dates/range to select</param>
        public void SelectRange(DateTime[] dates)
        {
            if (dates == null || dates.Length == 0)
            {
                throw new Exception("Invalid parameter");
            }
            
            if (uiElement.Current.FrameworkId == "WPF")
            {
                SetSelectedDate(dates[0], false);
                for (int i = 1; i < dates.Length; i++)
                {
                    SetSelectedDate(dates[i], true);
                }
            }
            else if (uiElement.Current.FrameworkId == "Win32")
            {
                SetSelectedRange(new IntPtr(uiElement.Current.NativeWindowHandle), dates);
            }
            else
            {
                throw new Exception("Not supported");
            }
        }
        
        /// <summary>
        /// Adds the specified date to selection. This method works only for WPF calendar.
        /// </summary>
        /// <param name="date">date to add to selection</param>
        public void AddToSelection(DateTime date)
        {
            SetSelectedDate(date, true);
        }
        
        /// <summary>
        /// Adds the specified range to selection. This method works only for WPF calendar.
        /// All dates of the range should be specified in the parameter array.
        /// </summary>
        /// <param name="dates">dates to add to selection</param>
        public void AddRangeToSelection(DateTime[] dates)
        {
            foreach (DateTime date in dates)
            {
                SetSelectedDate(date, true);
            }
        }
        
        private void SetSelectedDate(DateTime date, bool add)
        {
            object multipleViewPatternObj = null;
            uiElement.TryGetCurrentPattern(MultipleViewPattern.Pattern, out multipleViewPatternObj);
            MultipleViewPattern multipleViewPattern = multipleViewPatternObj as MultipleViewPattern;
            
            // switch to Decade view
            if (multipleViewPattern != null)
            {
                int[] views = multipleViewPattern.Current.GetSupportedViews();
                if (views.Length > 0)
                {
                    // the third view is the Decade view
                    multipleViewPattern.SetCurrentView(views[2]);
                }
            }
            
            // set year
            UIDA_Button btnHeader = this.Buttons()[1];
            if (btnHeader.InnerElement.Current.AutomationId == "PART_HeaderButton")
            {
                string headerName = btnHeader.InnerElement.Current.Name;
                string[] parts = headerName.Split('-');
                int yearLow = Convert.ToInt32(parts[0]);
                int yearHigh = Convert.ToInt32(parts[1]);
                if (date.Year < yearLow)
                {
                    UIDA_Button prevBtn = this.Buttons()[0];
                    while (date.Year < yearLow)
                    {
                        object invokePatternObj = null;
                        prevBtn.InnerElement.TryGetCurrentPattern(InvokePattern.Pattern, out invokePatternObj);
                        InvokePattern invokePattern = invokePatternObj as InvokePattern;
                        if (invokePattern == null)
                        {
                            break;
                        }
                        invokePattern.Invoke();
                        
                        headerName = btnHeader.InnerElement.Current.Name;
                        parts = headerName.Split('-');
                        yearLow = Convert.ToInt32(parts[0]);
                    }
                }
                else if (date.Year > yearHigh)
                {
                    UIDA_Button nextBtn = this.Buttons()[2];
                    while (date.Year > yearHigh)
                    {
                        object invokePatternObj = null;
                        nextBtn.InnerElement.TryGetCurrentPattern(InvokePattern.Pattern, out invokePatternObj);
                        InvokePattern invokePattern = invokePatternObj as InvokePattern;
                        if (invokePattern == null)
                        {
                            break;
                        }
                        invokePattern.Invoke();
                        //btnHeader = this.Buttons()[1];
                        headerName = btnHeader.InnerElement.Current.Name;
                        parts = headerName.Split('-');
                        yearHigh = Convert.ToInt32(parts[1]);
                    }
                }
                
                UIDA_Button[] buttons = this.Buttons();
                for (int i = 3; i < buttons.Length; i++)
                {
                    UIDA_Button button = buttons[i];
                    if (button.InnerElement.Current.Name == date.Year.ToString())
                    {
                        object invokePatternObj = null;
                        button.InnerElement.TryGetCurrentPattern(InvokePattern.Pattern, out invokePatternObj);
                        InvokePattern invokePattern = invokePatternObj as InvokePattern;
                        if (invokePattern != null)
                        {
                            invokePattern.Invoke();
                        }
                    }
                }
            }
            
            // set month
            UIDA_Button[] monthButtons = this.Buttons();
            for (int i = 3; i < monthButtons.Length; i++)
            {
                UIDA_Button monthBtn = monthButtons[i];
                DateTime crtMonthDate = DateTime.Parse(monthBtn.InnerElement.Current.Name, CultureInfo.CurrentCulture);
                
                if (crtMonthDate.Month == date.Month)
                {
                    object invokePatternObj = null;
                    monthBtn.InnerElement.TryGetCurrentPattern(InvokePattern.Pattern, out invokePatternObj);
                    InvokePattern invokePattern = invokePatternObj as InvokePattern;
                    if (invokePattern != null)
                    {
                        invokePattern.Invoke();
                    }
                }
            }
            
            // set day
            UIDA_Button[] dayButtons = this.Buttons();
            DateTime dateDayMonthYear = new DateTime(date.Year, date.Month, date.Day);
            
            for (int i = 3; i < dayButtons.Length; i++)
            {
                UIDA_Button dayBtn = dayButtons[i];
                string dayStr = dayBtn.InnerElement.Current.Name;
                DateTime currentDate;
				
                try
                {
                    currentDate = DateTime.Parse(dayStr, CultureInfo.CurrentCulture);
                }
                catch
                {
                    continue;
                }
                
                if (currentDate == dateDayMonthYear)
                {
                    if (add == true)
                    {
                        object selectionItemPatternObj = null;
                        dayBtn.InnerElement.TryGetCurrentPattern(SelectionItemPattern.Pattern, out selectionItemPatternObj);
                        SelectionItemPattern selectionItemPattern = selectionItemPatternObj as SelectionItemPattern;
                        if (selectionItemPattern != null)
                        {
                            selectionItemPattern.AddToSelection();
                        }
                    }
                    else
                    {
                        object invokePatternObj = null;
                        dayBtn.InnerElement.TryGetCurrentPattern(InvokePattern.Pattern, out invokePatternObj);
                        InvokePattern invokePattern = invokePatternObj as InvokePattern;
                        if (invokePattern != null)
                        {
                            invokePattern.Invoke();
                        }
                    }
                }
            }
        }
        
        private string GetWindowClassName(IntPtr handle)
        {
            if (handle == IntPtr.Zero)
            {
                return null;
            }
            StringBuilder className = new StringBuilder(256);
            UnsafeNativeFunctions.GetClassName(handle, className, 256);
            return className.ToString();
        }
        
        private DateTime[] GetSelection(IntPtr handle)
        {
            if (handle == IntPtr.Zero || GetWindowClassName(handle) != "SysMonthCal32")
            {
                throw new Exception("Not supported for this type of calendar");
            }
            
            uint styles = UnsafeNativeFunctions.GetWindowLong(handle, (int)GWL.GWL_STYLE);
            if ((styles & Win32CalendarStyles.MCS_MULTISELECT) != 0)
            {
                // multiple selection calendar
                DateTime[] dates = GetSelectedRange(handle);
                if (dates.Length == 2 && dates[0] == dates[1])
                {
                    return new DateTime[] { dates[0] };
                }
                return dates;
            }
            else
            {
                // single selection calendar
                DateTime date = GetSelectedDate(handle);
                return new DateTime[] { date };
            }
        }
        
        private DateTime[] GetSelectedRange(IntPtr handle)
        {
            uint procid = 0;
            UnsafeNativeFunctions.GetWindowThreadProcessId(handle, out procid);
            
            IntPtr hProcess = UnsafeNativeFunctions.OpenProcess(ProcessAccessFlags.All, false, (int)procid);
            if (hProcess == IntPtr.Zero)
            {
                throw new Exception("Insufficient rights");
            }
            
            SYSTEMTIME systemtime1 = new SYSTEMTIME();
            SYSTEMTIME systemtime2 = new SYSTEMTIME();
            IntPtr hMem = UnsafeNativeFunctions.VirtualAllocEx(hProcess, IntPtr.Zero, (uint)(2 * Marshal.SizeOf(systemtime1)),
                AllocationType.Commit | AllocationType.Reserve, MemoryProtection.ReadWrite);
            if (hMem == IntPtr.Zero)
            {
                throw new Exception("Insufficient rights");
            }
            
            UnsafeNativeFunctions.SendMessage(handle, Win32CalendarMessages.MCM_GETSELRANGE, IntPtr.Zero, hMem);
            
            IntPtr address = Marshal.AllocHGlobal(2 * Marshal.SizeOf(systemtime1));
            
            IntPtr lpNumberOfBytesRead = IntPtr.Zero;
            if (UnsafeNativeFunctions.ReadProcessMemory(hProcess, hMem, address, 2 * Marshal.SizeOf(systemtime1), 
                out lpNumberOfBytesRead) == false)
            {
                throw new Exception("Insufficient rights");
            }
                
            systemtime1 = (SYSTEMTIME)Marshal.PtrToStructure(address, typeof(SYSTEMTIME));
            IntPtr address2 = new IntPtr(address.ToInt64() + Marshal.SizeOf(systemtime1));
            systemtime2 = (SYSTEMTIME)Marshal.PtrToStructure(address2, typeof(SYSTEMTIME));
            
            Marshal.FreeHGlobal(address);
            UnsafeNativeFunctions.VirtualFreeEx(hProcess, hMem, 2 * Marshal.SizeOf(systemtime1), 
                FreeType.Decommit | FreeType.Release);
            UnsafeNativeFunctions.CloseHandle(hProcess);
            
            DateTime date1;
            try
            {
                date1 = new DateTime(systemtime1.Year, systemtime1.Month, systemtime1.Day, 
                    systemtime1.Hour, systemtime1.Minute, systemtime1.Second);
            }
            catch
            {
                date1 = new DateTime(systemtime1.Year, systemtime1.Month, systemtime1.Day, 
                    0, 0, 0);
            }
            
            Engine.TraceInLogFile("date1 = " + date1.ToString());
            
            DateTime date2;
            try
            {
                date2 = new DateTime(systemtime2.Year, systemtime2.Month, systemtime2.Day, 
                    systemtime2.Hour, systemtime2.Minute, systemtime2.Second);
            }
            catch
            {
                date2 = new DateTime(systemtime2.Year, systemtime2.Month, systemtime2.Day, 
                    0, 0, 0);
            }
            
            Engine.TraceInLogFile("date2 = " + date2.ToString());
            
            return new DateTime[] { date1, date2 };
        }
        
        private DateTime GetSelectedDate(IntPtr handle)
        {
            uint procid = 0;
            UnsafeNativeFunctions.GetWindowThreadProcessId(handle, out procid);
            
            IntPtr hProcess = UnsafeNativeFunctions.OpenProcess(ProcessAccessFlags.All, false, (int)procid);
            if (hProcess == IntPtr.Zero)
            {
                throw new Exception("Insufficient rights");
            }
            
            SYSTEMTIME systemtime = new SYSTEMTIME();
            IntPtr hMem = UnsafeNativeFunctions.VirtualAllocEx(hProcess, IntPtr.Zero, (uint)Marshal.SizeOf(systemtime), 
                AllocationType.Commit | AllocationType.Reserve, MemoryProtection.ReadWrite);
            if (hMem == IntPtr.Zero)
            {
                throw new Exception("Insufficient rights");
            }
            
            UnsafeNativeFunctions.SendMessage(handle, Win32CalendarMessages.MCM_GETCURSEL, IntPtr.Zero, hMem);
            
            IntPtr address = Marshal.AllocHGlobal(Marshal.SizeOf(systemtime));
            
            IntPtr lpNumberOfBytesRead = IntPtr.Zero;
            if (UnsafeNativeFunctions.ReadProcessMemory(hProcess, hMem, address, Marshal.SizeOf(systemtime), 
                out lpNumberOfBytesRead) == false)
            {
                throw new Exception("Insufficient rights");
            }

            systemtime = (SYSTEMTIME)Marshal.PtrToStructure(address, typeof(SYSTEMTIME));
            
            Marshal.FreeHGlobal(address);
            UnsafeNativeFunctions.VirtualFreeEx(hProcess, hMem, Marshal.SizeOf(systemtime), 
                FreeType.Decommit | FreeType.Release);
            UnsafeNativeFunctions.CloseHandle(hProcess);
            
            DateTime datetime;
            try
            {
                datetime = new DateTime(systemtime.Year, systemtime.Month, systemtime.Day, 
                    systemtime.Hour, systemtime.Minute, systemtime.Second);
            }
            catch
            {
                datetime = new DateTime(systemtime.Year, systemtime.Month, systemtime.Day, 
                    0, 0, 0);
            }
            
            return datetime;
        }
        
        private void SetSelectedDate(IntPtr handle, DateTime date)
        {
            if (handle == IntPtr.Zero || GetWindowClassName(handle) != "SysMonthCal32")
            {
                throw new Exception("Not supported for this type of calendar");
            }
            
            uint styles = UnsafeNativeFunctions.GetWindowLong(handle, (int)GWL.GWL_STYLE);
            if ((styles & Win32CalendarStyles.MCS_MULTISELECT) != 0)
            {
                // multiselect calendar
                SetSelectedRange(handle, new DateTime[] { date, date });
                return;
            }
            
            uint procid = 0;
            UnsafeNativeFunctions.GetWindowThreadProcessId(handle, out procid);
            
            IntPtr hProcess = UnsafeNativeFunctions.OpenProcess(ProcessAccessFlags.All, false, (int)procid);
            if (hProcess == IntPtr.Zero)
            {
                throw new Exception("Insufficient rights");
            }
            
            SYSTEMTIME systemtime = new SYSTEMTIME();
            systemtime.Year = (short)date.Year;
            systemtime.Month = (short)date.Month;
            systemtime.Day = (short)date.Day;
            systemtime.DayOfWeek = (short)date.DayOfWeek;
            systemtime.Hour = (short)date.Hour;
            systemtime.Minute = (short)date.Minute;
            systemtime.Second = (short)date.Second;
            systemtime.Milliseconds = (short)date.Millisecond;
            
            IntPtr hMem = UnsafeNativeFunctions.VirtualAllocEx(hProcess, IntPtr.Zero, (uint)Marshal.SizeOf(systemtime), 
                AllocationType.Commit | AllocationType.Reserve, MemoryProtection.ReadWrite);
            if (hMem == IntPtr.Zero)
            {
                throw new Exception("Insufficient rights");
            }
            
            IntPtr lpNumberOfBytesWritten = IntPtr.Zero;
            if (UnsafeNativeFunctions.WriteProcessMemory(hProcess, hMem, systemtime, Marshal.SizeOf(systemtime), 
                out lpNumberOfBytesWritten) == false)
            {
                throw new Exception("Insufficient rights");
            }
            
            UnsafeNativeFunctions.SendMessage(handle, Win32CalendarMessages.MCM_SETCURSEL, IntPtr.Zero, hMem);
            
            UnsafeNativeFunctions.VirtualFreeEx(hProcess, hMem, Marshal.SizeOf(systemtime), 
                FreeType.Decommit | FreeType.Release);
            UnsafeNativeFunctions.CloseHandle(hProcess);
        }
        
        private void SetSelectedRange(IntPtr handle, DateTime[] dates)
        {
            if (handle == IntPtr.Zero || GetWindowClassName(handle) != "SysMonthCal32")
            {
                throw new Exception("Not supported for this type of calendar");
            }
            
            if (dates.Length != 2)
            {
                throw new Exception("Dates array length must be 2");
            }
            
            uint styles = UnsafeNativeFunctions.GetWindowLong(handle, (int)GWL.GWL_STYLE);
            if ((styles & Win32CalendarStyles.MCS_MULTISELECT) == 0)
            {
                // singleselect calendar
                SetSelectedDate(handle, dates[1]);
                return;
            }
            
            uint procid = 0;
            UnsafeNativeFunctions.GetWindowThreadProcessId(handle, out procid);
            
            IntPtr hProcess = UnsafeNativeFunctions.OpenProcess(ProcessAccessFlags.All, false, (int)procid);
            if (hProcess == IntPtr.Zero)
            {
                throw new Exception("Insufficient rights");
            }
            
            SYSTEMTIME systemtime1 = new SYSTEMTIME();
            systemtime1.Year = (short)dates[0].Year;
            systemtime1.Month = (short)dates[0].Month;
            systemtime1.Day = (short)dates[0].Day;
            systemtime1.DayOfWeek = (short)dates[0].DayOfWeek;
            systemtime1.Hour = (short)dates[0].Hour;
            systemtime1.Minute = (short)dates[0].Minute;
            systemtime1.Second = (short)dates[0].Second;
            systemtime1.Milliseconds = (short)dates[0].Millisecond;
            
            SYSTEMTIME systemtime2 = new SYSTEMTIME();
            systemtime2.Year = (short)dates[1].Year;
            systemtime2.Month = (short)dates[1].Month;
            systemtime2.Day = (short)dates[1].Day;
            systemtime2.DayOfWeek = (short)dates[1].DayOfWeek;
            systemtime2.Hour = (short)dates[1].Hour;
            systemtime2.Minute = (short)dates[1].Minute;
            systemtime2.Second = (short)dates[1].Second;
            systemtime2.Milliseconds = (short)dates[1].Millisecond;
            
            IntPtr hMem = UnsafeNativeFunctions.VirtualAllocEx(hProcess, IntPtr.Zero, (uint)(2 * Marshal.SizeOf(systemtime1)),
                AllocationType.Commit | AllocationType.Reserve, MemoryProtection.ReadWrite);
            if (hMem == IntPtr.Zero)
            {
                throw new Exception("Insufficient rights");
            }
            
            IntPtr lpNumberOfBytesWritten = IntPtr.Zero;
            if (UnsafeNativeFunctions.WriteProcessMemory(hProcess, hMem, systemtime1, Marshal.SizeOf(systemtime1), 
                out lpNumberOfBytesWritten) == false)
            {
                throw new Exception("Insufficient rights");
            }
            IntPtr hMem2 = new IntPtr(hMem.ToInt64() + Marshal.SizeOf(systemtime1));
            if (UnsafeNativeFunctions.WriteProcessMemory(hProcess, hMem2, systemtime2, Marshal.SizeOf(systemtime2), 
                out lpNumberOfBytesWritten) == false)
            {
                throw new Exception("Insufficient rights");
            }
            
            UnsafeNativeFunctions.SendMessage(handle, Win32CalendarMessages.MCM_SETSELRANGE, IntPtr.Zero, hMem);
            
            UnsafeNativeFunctions.VirtualFreeEx(hProcess, hMem, 2 * Marshal.SizeOf(systemtime1),
                FreeType.Decommit | FreeType.Release);
            UnsafeNativeFunctions.CloseHandle(hProcess);
        }
    }
}
