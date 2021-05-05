using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using System.Globalization;
using System.Runtime.InteropServices;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Represents a DatePicker UI element
    /// </summary>
    public class DatePicker: ElementBase
    {
        public DatePicker(AutomationElement el)
        {
            base.uiElement = el;
        }
        
        /// <summary>
        /// Gets and sets selected date in the DatePicker control.
        /// For Win32 DateTimePicker setting the SelectedDate to null will uncheck the control and disable it. If a Win32 DateTimePicker is unchecked then SelectedDate will return null.
        /// </summary>
        /// <returns>DateTime selected in the DatePicker</returns>
        public DateTime? SelectedDate
        {
            get
            {
                string fid = this.uiElement.Current.FrameworkId;
                if (fid == "WPF")
                {
                    object valuePatternObj = null;
                    if (this.uiElement.TryGetCurrentPattern(ValuePatternIdentifiers.Pattern, 
                        out valuePatternObj) == true)
                    {
                        ValuePattern valuePattern = valuePatternObj as ValuePattern;
                        if (valuePattern != null)
                        {
                            try
                            {
                                string val = valuePattern.Current.Value;
                                DateTime date = DateTime.Parse(val, CultureInfo.CurrentCulture);
                                return date;
                            }
                            catch { }
                        }
                    }
                }
                else if (fid == "Win32")
                {
                    IntPtr handle = new IntPtr(this.uiElement.Current.NativeWindowHandle);
                    return GetSelectedDate(handle);
                }
                else if (fid == "WinForm")
                {
                    string name = this.uiElement.Current.Name;
                    string dateString = name.Remove(0, name.IndexOf(",") + 1).Trim();
                    
                    DateTime date = DateTime.Parse(dateString, CultureInfo.CurrentCulture);
                    return date;
                }
                return null;
            }
            set
            {
                string fid = this.uiElement.Current.FrameworkId;
                if (fid == "WPF")
                {
                    object valuePatternObj = null;
                    if (this.uiElement.TryGetCurrentPattern(ValuePatternIdentifiers.Pattern, 
                        out valuePatternObj) == true)
                    {
                        ValuePattern valuePattern = valuePatternObj as ValuePattern;
                        if (valuePattern != null)
                        {
                            try
                            {
                                if (value.HasValue)
                                {
                                    valuePattern.SetValue(value.Value.ToString(CultureInfo.CurrentCulture));
                                }
                            }
                            catch { }
                        }
                    }
                }
                else if (fid == "Win32")
                {
                    IntPtr handle = new IntPtr(this.uiElement.Current.NativeWindowHandle);
                    SetSelectedDate(handle, value);
                }
                else if (fid == "WinForm")
                {
                    if (value.HasValue == false)
                    {
                        return;
                    }
                    
                    System.Windows.Rect boundingRect = this.uiElement.Current.BoundingRectangle;
                    int x = (int)boundingRect.Right - 5;
                    int y = (int)((boundingRect.Top + boundingRect.Bottom) / 2);
                    
                    Engine.GetInstance().ClickScreenCoordinatesAt(x, y);
                    
                    TreeWalker tw = TreeWalker.ControlViewWalker;
                    AutomationElement parent = tw.GetParent(this.uiElement);
                    AutomationElement parentParent = tw.GetParent(parent);
                    AutomationElement root = AutomationElement.RootElement;
                    
                    while (parentParent != root)
                    {
                        parent = parentParent;
                        parentParent = tw.GetParent(parent);
                    }
                    
                    Window window = new Window(parent);
                    Calendar calendar = window.Calendar("Calendar Control", true);
                    
                    if (calendar != null)
                    {
                        calendar.SelectDate(value.Value);
                        
                        SendKeys(" ");
                    }
                }
            }
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
            
            UnsafeNativeFunctions.SendMessage(handle, DateTimePicker32Messages.DTM_GETSYSTEMTIME, IntPtr.Zero, hMem);
            
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
        
        private void SetSelectedDate(IntPtr handle, DateTime? date)
        {
            if (handle == IntPtr.Zero || GetWindowClassName(handle) != "SysDateTimePick32")
            {
                throw new Exception("Not supported for this type of DateTimePicker");
            }
            
            if (date == null)
            {
                UnsafeNativeFunctions.SendMessage(handle, DateTimePicker32Messages.DTM_SETSYSTEMTIME, new IntPtr(DateTimePicker32Constants.GDT_NONE), IntPtr.Zero);
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
            systemtime.Year = (short)date.Value.Year;
            systemtime.Month = (short)date.Value.Month;
            systemtime.Day = (short)date.Value.Day;
            systemtime.DayOfWeek = (short)date.Value.DayOfWeek;
            systemtime.Hour = (short)date.Value.Hour;
            systemtime.Minute = (short)date.Value.Minute;
            systemtime.Second = (short)date.Value.Second;
            systemtime.Milliseconds = (short)date.Value.Millisecond;
            
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
            
            UnsafeNativeFunctions.SendMessage(handle, DateTimePicker32Messages.DTM_SETSYSTEMTIME, new IntPtr(DateTimePicker32Constants.GDT_VALID), hMem);
            
            UnsafeNativeFunctions.VirtualFreeEx(hProcess, hMem, Marshal.SizeOf(systemtime), 
                FreeType.Decommit | FreeType.Release);
            UnsafeNativeFunctions.CloseHandle(hProcess);
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
    }
}
