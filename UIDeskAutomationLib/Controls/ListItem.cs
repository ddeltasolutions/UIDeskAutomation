using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a listitem UI element.
    /// </summary>
    public class UIDA_ListItem: ElementBase
    {
		/// <summary>
        /// Creates a UIDA_ListItem using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_ListItem(AutomationElement el)
        {
            base.uiElement = el;
        }

        /// <summary>
        /// Deselects any selected items and selects the current list item
        /// </summary>
        public void Select()
        { 
            object selectionItemPatternObj = null;
            if (this.uiElement.TryGetCurrentPattern(SelectionItemPattern.Pattern,
                out selectionItemPatternObj) == true)
            {
                SelectionItemPattern selectionItemPattern =
                    selectionItemPatternObj as SelectionItemPattern;

                if (selectionItemPattern != null)
                {
                    try
                    {
                        selectionItemPattern.Select();
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile("ListItem.Select failed: " + ex.Message);
                    }
                }
            }
        }

        /// <summary>
        /// Removes the current list items from the collection of selected list items
        /// </summary>
        public void RemoveFromSelection()
        {
            object selectionItemPatternObj = null;
            if (this.uiElement.TryGetCurrentPattern(SelectionItemPattern.Pattern,
                out selectionItemPatternObj) == true)
            {
                SelectionItemPattern selectionItemPattern =
                    selectionItemPatternObj as SelectionItemPattern;

                if (selectionItemPattern != null)
                {
                    try
                    {
                        selectionItemPattern.RemoveFromSelection();
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile("ListItem.Deselect failed: " + ex.Message);
                    }
                }
            }
        }

        /// <summary>
        /// Adds the current list item to the collection of selected list items
        /// </summary>
        public void AddToSelection()
        {
            object selectionItemPatternObj = null;
            if (this.uiElement.TryGetCurrentPattern(SelectionItemPattern.Pattern,
                out selectionItemPatternObj) == true)
            {
                SelectionItemPattern selectionItemPattern =
                    selectionItemPatternObj as SelectionItemPattern;

                if (selectionItemPattern != null)
                {
                    try
                    {
                        selectionItemPattern.AddToSelection();
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile("AddToSelection failed: " + ex.Message);
                    }
                }
            }
        }

        /// <summary>
        /// Gets the zero based index of this list item in the list items collection.
        /// </summary>
        public int Index
        {
            get
            {
                TreeWalker treeWalker = TreeWalker.ControlViewWalker;
                AutomationElement parent = treeWalker.GetParent(this.uiElement);

                while (parent != null)
                {
                    try
                    {
                        if (parent.Current.ControlType == ControlType.List)
                        {
                            break;
                        }
                    }
                    catch
                    {
                        break;
                    }

                    parent = treeWalker.GetParent(parent);
                }

                if (parent == null)
                {
                    Engine.TraceInLogFile("Error getting ListItem index");
                    throw new Exception("Error getting ListItem index");
                }

                UIDA_List parentList = new UIDA_List(parent);
                UIDA_ListItem[] listItems = parentList.Items;

                int index = -1;

                foreach (UIDA_ListItem currentListItem in listItems)
                {
                    index++;

                    int[] runtimeId1 = null;
                    int[] runtimeId2 = null;

                    try
                    {
                        runtimeId1 = this.uiElement.GetRuntimeId();
                        runtimeId2 = currentListItem.uiElement.GetRuntimeId();
                    }
                    catch { }

                    if ((runtimeId1 == null) || (runtimeId2 == null))
                    {
                        continue;
                    }

                    if (runtimeId1.Length != runtimeId2.Length)
                    {
                        continue;
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

                    if (different == true)
                    {
                        continue;
                    }

                    return index;
                }

                return -1; // not found.
            }
        }

        /// <summary>
        /// Returns true if the current list item is selected, false otherwise.
        /// </summary>
        public bool IsSelected
        {
            get
            {
                object selectionItemPatternObj = null;
                if (this.uiElement.TryGetCurrentPattern(SelectionItemPattern.Pattern,
                    out selectionItemPatternObj) == true)
                {
                    SelectionItemPattern selectionItemPattern =
                        selectionItemPatternObj as SelectionItemPattern;

                    if (selectionItemPattern != null)
                    {
                        try
                        {
                            return selectionItemPattern.Current.IsSelected;
                        }
                        catch (Exception ex)
                        {
                            Engine.TraceInLogFile("ListItem.IsSelected failed: " +
                                ex.Message);
                        }
                    }
                }

                return false;
            }
        }

        /// <summary>
        /// Brings the current List item into viewable area of the parent List control.
        /// </summary>
        public void BringIntoView()
        {
            object scrollItemPatternObj = null;

            if (this.uiElement.TryGetCurrentPattern(ScrollItemPattern.Pattern,
                out scrollItemPatternObj) == true)
            {
                ScrollItemPattern scrollItemPattern =
                    scrollItemPatternObj as ScrollItemPattern;

                if (scrollItemPattern == null)
                {
                    Engine.TraceInLogFile("ListItem.BringIntoView method failed");
                    throw new Exception("ListItem.BringIntoView method failed");
                }

                try
                {
                    scrollItemPattern.ScrollIntoView();

                    return;
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("ListItem.BringIntoView method failed: " +
                        ex.Message);

                    throw new Exception("ListItem.BringIntoView method failed: " +
                        ex.Message);
                }
            }

            Engine.TraceInLogFile("ListItem.BringIntoView method failed");
            throw new Exception("ListItem.BringIntoView method failed");
        }

        /// <summary>
        /// Gets or sets the checked state of the current list item.
		/// true = checked, false = unchecked, null = indeterminate
        /// </summary>
        public bool? IsChecked
        {
            get
            {
                IntPtr hwndList = this.GetWindow();

                if (hwndList != IntPtr.Zero)
                {
                    StringBuilder className = new StringBuilder(256);
                    UnsafeNativeFunctions.GetClassName(hwndList, className, 256);

                    if (className.ToString() == "SysListView32")
                    { 
                        // Win32 standard listview control
                        int index = this.Index;

                        IntPtr itemState = UnsafeNativeFunctions.SendMessage(
				            hwndList, WindowMessages.LVM_GETITEMSTATE, 
                            new IntPtr(index), 
                            new IntPtr(Win32Constants.LVIS_STATEIMAGEMASK));

                        return ((itemState.ToInt32() & 0x2000) != 0);
                    }
                }

                object togglePatternObj = null;

                if (this.uiElement.TryGetCurrentPattern(TogglePattern.Pattern,
                    out togglePatternObj) == true)
                {
                    TogglePattern togglePattern = togglePatternObj as TogglePattern;

                    if (togglePattern == null)
                    {
                        Engine.TraceInLogFile("ListItem.IsChecked failed");
                        throw new Exception("ListItem.IsChecked failed");
                    }

                    try
                    {
                        if (togglePattern.Current.ToggleState == ToggleState.On)
                        {
                            return true;
                        }
                        else if (togglePattern.Current.ToggleState == ToggleState.Off)
                        {
                            return false;
                        }
						else
						{
							// indeterminate
							return null;
						}
                    }
                    catch (Exception ex)
                    { 
                        Engine.TraceInLogFile("ListItem.IsChecked failed: " + ex.Message);
                        throw new Exception("ListItem.IsChecked failed: " + ex.Message);
                    }
                }

                Engine.TraceInLogFile("ListItem.IsChecked failed");
                throw new Exception("ListItem.IsChecked failed");
            }

            set
            {
                IntPtr hwndList = this.GetWindow();

                if (hwndList != IntPtr.Zero)
                {
                    StringBuilder className = new StringBuilder(256);
                    UnsafeNativeFunctions.GetClassName(hwndList, className, 256);

                    if (className.ToString() == "SysListView32")
                    {
                        LV_ITEM lvItem = new LV_ITEM();

                        int index = this.Index;

                        lvItem.stateMask = Win32Constants.LVIS_STATEIMAGEMASK;
                        if (value == true)
                        {
                            lvItem.state = 0x2000; // check list item
                        }
                        else
                        {
                            lvItem.state = 0x1000; // uncheck list item
                        }

                        uint processId = 0;

                        UnsafeNativeFunctions.GetWindowThreadProcessId(hwndList,
                            out processId);

                        IntPtr hProcess = UnsafeNativeFunctions.OpenProcess(
                            ProcessAccessFlags.VirtualMemoryWrite | ProcessAccessFlags.VirtualMemoryOperation, 
                            false, (int)processId);

                        if (hProcess != IntPtr.Zero)
                        {
                            IntPtr lvItemPtrExtern = UnsafeNativeFunctions.VirtualAllocEx(
                                hProcess, IntPtr.Zero, (uint)Marshal.SizeOf(lvItem), 
                                AllocationType.Reserve | AllocationType.Commit, 
                                MemoryProtection.ReadWrite);
                            Debug.Assert(lvItemPtrExtern != IntPtr.Zero);
                            
                            IntPtr lvItemPtr = Marshal.AllocHGlobal(Marshal.SizeOf(lvItem));
                            Debug.Assert(lvItemPtr != IntPtr.Zero);
                            
                            System.Runtime.InteropServices.Marshal.StructureToPtr(lvItem,
                                lvItemPtr, false);

                            IntPtr numberOfBytesWritten = IntPtr.Zero;

                            bool bResult = UnsafeNativeFunctions.WriteProcessMemory(
                                hProcess, lvItemPtrExtern, lvItemPtr,
                                Marshal.SizeOf(lvItem), out numberOfBytesWritten);
                            Debug.Assert(bResult == true);

                            Marshal.FreeHGlobal(lvItemPtr);

                            IntPtr retPtr = UnsafeNativeFunctions.SendMessage(hwndList,
                                WindowMessages.LVM_SETITEMSTATE, new IntPtr(index), lvItemPtrExtern);
                            Debug.Assert(retPtr != IntPtr.Zero);

                            bResult = UnsafeNativeFunctions.VirtualFreeEx(hProcess, lvItemPtrExtern,
                                Marshal.SizeOf(lvItem), FreeType.Decommit | FreeType.Release);
                            //Debug.Assert(bResult == true);

                            bResult = UnsafeNativeFunctions.CloseHandle(hProcess);
                            Debug.Assert(bResult == true);
                        }

                        return;
                    }
                }

                object togglePatternObj = null;

                if (this.uiElement.TryGetCurrentPattern(TogglePattern.Pattern,
                    out togglePatternObj) == true)
                {
                    TogglePattern togglePattern = togglePatternObj as TogglePattern;

                    if (togglePattern == null)
                    {
                        if (value == true)
                        {
                            Engine.TraceInLogFile("Cannot check list item");
                            throw new Exception("Cannot check list item");
                        }
                        else if (value == false)
                        {
                            Engine.TraceInLogFile("Cannot uncheck list item");
                            throw new Exception("Cannot uncheck list item");
                        }
						else
						{
							Engine.TraceInLogFile("Cannot set the list item checked state");
                            throw new Exception("Cannot set the list item checked state");
						}
                    }

                    if (value == true)
                    {
                        // try to check list item
                        try
                        {
                            if (togglePattern.Current.ToggleState != ToggleState.On)
                            {
                                togglePattern.Toggle();
                            }
							
							if (togglePattern.Current.ToggleState != ToggleState.On)
                            {
                                togglePattern.Toggle();
                            }

                            if (togglePattern.Current.ToggleState != ToggleState.On)
                            {
                                //this.SimulateDoubleClick();
                                this.DoubleClick();
                            }

                            return;
                        }
                        catch (Exception ex)
                        {
                            Engine.TraceInLogFile("Cannot check list item: " + ex.Message);
                            throw new Exception("Cannot check list item: " + ex.Message);
                        }
                    }
                    else if (value == false)
                    {
                        // try to uncheck
                        try
                        {
                            if (togglePattern.Current.ToggleState != ToggleState.Off)
                            {
                                togglePattern.Toggle();
                            }
							
							if (togglePattern.Current.ToggleState != ToggleState.Off)
                            {
                                togglePattern.Toggle();
                            }

                            if (togglePattern.Current.ToggleState != ToggleState.Off)
                            {
                                //this.SimulateDoubleClick();
								this.DoubleClick();
                            }

                            return;
                        }
                        catch (Exception ex)
                        {
                            Engine.TraceInLogFile("Cannot uncheck list item: " + ex.Message);
                            throw new Exception("Cannot uncheck list item: " + ex.Message);
                        }
                    }
					else // indeterminate
					{
                        try
                        {
                            if (togglePattern.Current.ToggleState != ToggleState.Indeterminate)
                            {
                                togglePattern.Toggle();
                            }
							
							if (togglePattern.Current.ToggleState != ToggleState.Indeterminate)
                            {
                                togglePattern.Toggle();
                            }

                            if (togglePattern.Current.ToggleState != ToggleState.Indeterminate)
                            {
                                //this.SimulateDoubleClick();
                                this.DoubleClick();
                            }
							if (togglePattern.Current.ToggleState != ToggleState.Indeterminate)
                            {
                                this.DoubleClick();
                            }

                            return;
                        }
                        catch (Exception ex)
                        {
                            Engine.TraceInLogFile("Cannot check list item: " + ex.Message);
                            throw new Exception("Cannot check list item: " + ex.Message);
                        }
					}
                }

                if (value == true)
                {
                    Engine.TraceInLogFile("Cannot check list item");
                    throw new Exception("Cannot check list item");
                }
                else if (value == false)
                { 
                    Engine.TraceInLogFile("Cannot uncheck list item");
                    throw new Exception("Cannot uncheck list item");
                }
				else
				{
					Engine.TraceInLogFile("Cannot set the checked state of the list item");
                    throw new Exception("Cannot set the checked state of the list item");
				}
            }
        }
    }
}
