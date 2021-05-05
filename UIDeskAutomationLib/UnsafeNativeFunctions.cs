using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace dDeltaSolutions.PSLib
{
    internal class UnsafeNativeFunctions
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent,
            IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, WindowShowStyle nCmdShow);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "SendMessage", CharSet = CharSet.Auto)]
        public static extern bool SendMessage(IntPtr hWnd, uint Msg, int wParam, StringBuilder lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("User32.dll")]
        public static extern Int32 SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern IntPtr GetFocus();

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool SetWindowText(IntPtr hwnd, String lpString);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool ScreenToClient(IntPtr hwnd, ref POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(POINT Point);

        [DllImport("user32.dll")]
        public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo,
           bool fAttach);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        public static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("kernel32.dll")]
        public static extern IntPtr OpenProcess(
             ProcessAccessFlags processAccess,
             bool bInheritHandle,
             int processId
        );

        [DllImport("kernel32.dll")]
        public static extern int GetProcessId(IntPtr processHandle);

        [DllImport("psapi.dll")]
        public static extern uint GetModuleFileNameEx(IntPtr hProcess, IntPtr hModule, 
            [Out] StringBuilder lpBaseName, [In] [MarshalAs(UnmanagedType.U4)] int nSize);

        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        public static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress,
           uint dwSize, AllocationType flAllocationType, MemoryProtection flProtect);

        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        public static extern bool VirtualFreeEx(IntPtr hProcess, IntPtr lpAddress,
           int dwSize, FreeType dwFreeType);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool CloseHandle(IntPtr hObject);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool WriteProcessMemory(
            IntPtr hProcess,
            IntPtr lpBaseAddress,
            byte[] lpBuffer,
            int nSize,
            out IntPtr lpNumberOfBytesWritten);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool WriteProcessMemory(
            IntPtr hProcess,
            IntPtr lpBaseAddress,
            IntPtr lpBuffer,
            int nSize,
            out IntPtr lpNumberOfBytesWritten);
            
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress,
            [MarshalAs(UnmanagedType.AsAny)] object lpBuffer, int dwSize,
            out IntPtr lpNumberOfBytesWritten);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);
        
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress,
            IntPtr lpBuffer, int dwSize, out IntPtr lpNumberOfBytesRead);
            
        [DllImport("user32.dll", EntryPoint="GetWindowLong")]
        public static extern uint GetWindowLong(IntPtr hWnd, int nIndex);
        
        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowDC(IntPtr hWnd);
        
        [DllImport("gdi32.dll", EntryPoint = "CreateCompatibleDC", SetLastError=true)]
        public static extern IntPtr CreateCompatibleDC([In] IntPtr hdc);
        
        [DllImport("gdi32.dll", EntryPoint = "CreateCompatibleBitmap")]
        public static extern IntPtr CreateCompatibleBitmap([In] IntPtr hdc, int nWidth, int nHeight);
        
        [DllImport("gdi32.dll", EntryPoint = "SelectObject")]
        public static extern IntPtr SelectObject([In] IntPtr hdc, [In] IntPtr hgdiobj);
        
        [DllImport("gdi32.dll", EntryPoint = "BitBlt", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool BitBlt([In] IntPtr hdc, int nXDest, int nYDest, int nWidth, int nHeight, 
            [In] IntPtr hdcSrc, int nXSrc, int nYSrc, TernaryRasterOperations dwRop);
            
        [DllImport("gdi32.dll", EntryPoint = "DeleteObject")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DeleteObject([In] IntPtr hObject);
        
        [DllImport("gdi32.dll", EntryPoint = "DeleteDC")]
        public static extern bool DeleteDC([In] IntPtr hdc);
        
        [DllImport("user32.dll")]
        public static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);
        
        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hWnd);
    }

    internal class ButtonMessages
    {
        public static UInt32 BM_SETCHECK = 0x00F1;  // set radio or checkbox button state
        public static UInt32 BM_GETCHECK = 0x00F0;

        public static UInt32 BST_UNCHECKED = 0x0000;
        public static UInt32 BST_CHECKED = 0x0001;
		public static UInt32 BST_INDETERMINATE = 0x0002;
    }

    internal class WindowMessages
    {
        public static UInt32 WM_SETTEXT = 0x000C;
        public static UInt32 WM_GETTEXT = 0x000D;
        public static UInt32 WM_GETTEXTLENGTH = 0x000E;
        public static UInt32 WM_CHAR = 0x0102;
        public static UInt32 WM_LBUTTONDOWN = 0x0201;
        public static UInt32 WM_LBUTTONUP = 0x0202;
        public static UInt32 WM_CLOSE = 0x0010;
        public static UInt32 WM_MBUTTONDOWN = 0x0207;
        public static UInt32 WM_MBUTTONUP = 0x0208;
        public static UInt32 WM_RBUTTONDOWN = 0x0204;
        public static UInt32 WM_RBUTTONUP = 0x0205;
        public static UInt32 WM_LBUTTONDBLCLK = 0x0203;
        public static UInt32 WM_KEYDOWN = 0x0100;
        public static UInt32 WM_KEYUP = 0x0101;
        public static UInt32 LVM_FIRST = 0x1000;
        public static UInt32 LVM_SETITEMSTATE = LVM_FIRST + 43;
        public static UInt32 LVM_GETITEMSTATE = LVM_FIRST + 44;
    }
    
    internal enum TernaryRasterOperations : uint 
    {
        SRCCOPY     = 0x00CC0020,
        SRCPAINT    = 0x00EE0086,
        SRCAND      = 0x008800C6,
        SRCINVERT   = 0x00660046,
        SRCERASE    = 0x00440328,
        NOTSRCCOPY  = 0x00330008,
        NOTSRCERASE = 0x001100A6,
        MERGECOPY   = 0x00C000CA,
        MERGEPAINT  = 0x00BB0226,
        PATCOPY     = 0x00F00021,
        PATPAINT    = 0x00FB0A09,
        PATINVERT   = 0x005A0049,
        DSTINVERT   = 0x00550009,
        BLACKNESS   = 0x00000042,
        WHITENESS   = 0x00FF0062,
        CAPTUREBLT  = 0x40000000 //only if WinVer >= 5.0.0 (see wingdi.h)
    }

    internal class VirtualKeyCodes
    {
        /// <summary>
        /// Control-break
        /// </summary>
        public static UInt32 VK_CANCEL = 0x03;
        /// <summary>
        /// Backspace
        /// </summary>
        public static UInt32 VK_BACK = 0x08;
        public static UInt32 VK_TAB = 0x09;
        public static UInt32 VK_CLEAR = 0x0C;
        /// <summary>
        /// Enter key
        /// </summary>
        public static UInt32 VK_RETURN = 0x0D;
        public static UInt32 VK_SHIFT = 0x10;
        public static UInt32 VK_CONTROL = 0x11;
        /// <summary>
        /// Alt key
        /// </summary>
        public static UInt32 VK_MENU = 0x12;
        public static UInt32 VK_PAUSE = 0x13;
        /// <summary>
        /// Caps lock key
        /// </summary>
        public static UInt32 VK_CAPITAL = 0x14; // CAPS LOCK
        /// <summary>
        /// ESC key
        /// </summary>
        public static UInt32 VK_ESCAPE = 0x1B; // ESC key
        public static UInt32 VK_SPACE = 0x20; // space
        /// <summary>
        /// Page UP key
        /// </summary>
        public static UInt32 VK_PRIOR = 0x21; // page up
        /// <summary>
        /// PAGE DOWN key
        /// </summary>
        public static UInt32 VK_NEXT = 0x22;
        public static UInt32 VK_END = 0x23;
        public static UInt32 VK_HOME = 0x24;
        /// <summary>
        /// Left arrow key
        /// </summary>
        public static UInt32 VK_LEFT = 0x25;
        /// <summary>
        /// Up arrow key
        /// </summary>
        public static UInt32 VK_UP = 0x26;
        public static UInt32 VK_RIGHT = 0x27;
        public static UInt32 VK_DOWN = 0x28;
        public static UInt32 VK_SNAPSHOT = 0x2C;
        public static UInt32 VK_INSERT = 0x2D;
        /// <summary>
        /// Del key
        /// </summary>
        public static UInt32 VK_DELETE = 0x2E;
        public static UInt32 VK_HELP = 0x2F;

        public static UInt32 VK_A = 0x41; // A key
        public static UInt32 VK_B = 0x42; // B key
        public static UInt32 VK_C = 0x43; // C key

        public static UInt32 VK_F1 = 0x70;
        public static UInt32 VK_F2 = 0x71;
        public static UInt32 VK_F3 = 0x72;
        public static UInt32 VK_F4 = 0x73;
        public static UInt32 VK_F5 = 0x74;
        public static UInt32 VK_F6 = 0x75;
        public static UInt32 VK_F7 = 0x76;
        public static UInt32 VK_F8 = 0x77;
        public static UInt32 VK_F9 = 0x78;
        public static UInt32 VK_F10 = 0x79;
        public static UInt32 VK_F11 = 0x7A;
        public static UInt32 VK_F12 = 0x7B;
        public static UInt32 VK_F13 = 0x7C;
        public static UInt32 VK_F14 = 0x7D;
        public static UInt32 VK_F15 = 0x7E;
        public static UInt32 VK_F16 = 0x7F;

        public static UInt32 VK_ADD = 0x6B;
        public static UInt32 VK_SUBTRACT = 0x6D;
        public static UInt32 VK_MULTIPLY = 0x6A;
        public static UInt32 VK_DIVIDE = 0x6F;
    }

    internal class Win32Constants
    {
        public static IntPtr TRUE = new IntPtr(1);
        public static IntPtr FALSE = new IntPtr(0);
        
        /// <summary>
        /// Left mouse button
        /// </summary>
        public static int MK_LBUTTON = 0x0001;
        /// <summary>
        /// Middle mouse button
        /// </summary>
        public static int MK_MBUTTON = 0x0010;
        /// <summary>
        /// Right mouse button
        /// </summary>
        public static int MK_RBUTTON = 0x0002;
        /// <summary>
        /// Control key is pressed
        /// </summary>
        public static int MK_CONTROL = 0x0008;
        /// <summary>
        /// Shift key is down
        /// </summary>
        public static int MK_SHIFT = 0x0004;

        public static uint LVIS_STATEIMAGEMASK = 0xF000;

        public static int PROCESS_VM_WRITE = 0x0020;
        public static int PROCESS_VM_OPERATION = 0x0008;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct LV_ITEM
    {
        public UInt32 mask;
        public Int32 iItem;
        public Int32 iSubItem;
        public UInt32 state;
        public UInt32 stateMask;
        public IntPtr pszText;
        public Int32 cchTextMax;
        public Int32 iImage;
        public IntPtr lParam;
        public Int32 iIndent;
        public Int32 iGroupId;
        public UInt32 cColumns;
        public UIntPtr puColumns;
        public IntPtr piColFmt;
        public Int32 iGroup;
    }

    enum WindowShowStyle : uint
    {
        /// <summary>Hides the window and activates another window.</summary>
        /// <remarks>See SW_HIDE</remarks>
        Hide = 0,
        /// <summary>Activates and displays a window. If the window is minimized 
        /// or maximized, the system restores it to its original size and 
        /// position. An application should specify this flag when displaying 
        /// the window for the first time.</summary>
        /// <remarks>See SW_SHOWNORMAL</remarks>
        ShowNormal = 1,
        /// <summary>Activates the window and displays it as a minimized window.</summary>
        /// <remarks>See SW_SHOWMINIMIZED</remarks>
        ShowMinimized = 2,
        /// <summary>Activates the window and displays it as a maximized window.</summary>
        /// <remarks>See SW_SHOWMAXIMIZED</remarks>
        ShowMaximized = 3,
        /// <summary>Maximizes the specified window.</summary>
        /// <remarks>See SW_MAXIMIZE</remarks>
        Maximize = 3,
        /// <summary>Displays a window in its most recent size and position. 
        /// This value is similar to "ShowNormal", except the window is not 
        /// actived.</summary>
        /// <remarks>See SW_SHOWNOACTIVATE</remarks>
        ShowNormalNoActivate = 4,
        /// <summary>Activates the window and displays it in its current size 
        /// and position.</summary>
        /// <remarks>See SW_SHOW</remarks>
        Show = 5,
        /// <summary>Minimizes the specified window and activates the next 
        /// top-level window in the Z order.</summary>
        /// <remarks>See SW_MINIMIZE</remarks>
        Minimize = 6,
        /// <summary>Displays the window as a minimized window. This value is 
        /// similar to "ShowMinimized", except the window is not activated.</summary>
        /// <remarks>See SW_SHOWMINNOACTIVE</remarks>
        ShowMinNoActivate = 7,
        /// <summary>Displays the window in its current size and position. This 
        /// value is similar to "Show", except the window is not activated.</summary>
        /// <remarks>See SW_SHOWNA</remarks>
        ShowNoActivate = 8,
        /// <summary>Activates and displays the window. If the window is 
        /// minimized or maximized, the system restores it to its original size 
        /// and position. An application should specify this flag when restoring 
        /// a minimized window.</summary>
        /// <remarks>See SW_RESTORE</remarks>
        Restore = 9,
        /// <summary>Sets the show state based on the SW_ value specified in the 
        /// STARTUPINFO structure passed to the CreateProcess function by the 
        /// program that started the application.</summary>
        /// <remarks>See SW_SHOWDEFAULT</remarks>
        ShowDefault = 10,
        /// <summary>Windows 2000/XP: Minimizes a window, even if the thread 
        /// that owns the window is hung. This flag should only be used when 
        /// minimizing windows from a different thread.</summary>
        /// <remarks>See SW_FORCEMINIMIZE</remarks>
        ForceMinimized = 11
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct RECT
    {
        public int Left;        // x position of upper-left corner
        public int Top;         // y position of upper-left corner
        public int Right;       // x position of lower-right corner
        public int Bottom;      // y position of lower-right corner

        public RECT(int left, int top, int right, int bottom)
        {
            this.Left = left;
            this.Top = top;
            this.Right = right;
            this.Bottom = bottom;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct POINT
    {
        public int X;
        public int Y;

        public POINT(int x, int y)
        {
            this.X = x;
            this.Y = y;
        }

        public POINT(System.Drawing.Point pt) : this(pt.X, pt.Y) { }

        public static implicit operator System.Drawing.Point(POINT p)
        {
            return new System.Drawing.Point(p.X, p.Y);
        }

        public static implicit operator POINT(System.Drawing.Point p)
        {
            return new POINT(p.X, p.Y);
        }
    }

    [Flags]
    internal enum ProcessAccessFlags : uint
    {
        All = 0x001F0FFF,
        Terminate = 0x00000001,
        CreateThread = 0x00000002,
        VirtualMemoryOperation = 0x00000008,
        VirtualMemoryRead = 0x00000010,
        VirtualMemoryWrite = 0x00000020,
        DuplicateHandle = 0x00000040,
        CreateProcess = 0x000000080,
        SetQuota = 0x00000100,
        SetInformation = 0x00000200,
        QueryInformation = 0x00000400,
        QueryLimitedInformation = 0x00001000,
        Synchronize = 0x00100000
    }

    [Flags]
    internal enum AllocationType
    {
        Commit = 0x1000,
        Reserve = 0x2000,
        Decommit = 0x4000,
        Release = 0x8000,
        Reset = 0x80000,
        Physical = 0x400000,
        TopDown = 0x100000,
        WriteWatch = 0x200000,
        LargePages = 0x20000000
    }

    [Flags]
    internal enum MemoryProtection
    {
        Execute = 0x10,
        ExecuteRead = 0x20,
        ExecuteReadWrite = 0x40,
        ExecuteWriteCopy = 0x80,
        NoAccess = 0x01,
        ReadOnly = 0x02,
        ReadWrite = 0x04,
        WriteCopy = 0x08,
        GuardModifierflag = 0x100,
        NoCacheModifierflag = 0x200,
        WriteCombineModifierflag = 0x400
    }

    [Flags]
    internal enum FreeType
    {
        Decommit = 0x4000,
        Release = 0x8000,
    }

    internal class SendInputClass
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(uint nInputs, ref INPUT pInputs, int cbSize);

        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT
        {
            public SendInputEventType type;
            public MouseKeybdhardwareInputUnion mkhi;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct MouseKeybdhardwareInputUnion
        {
            [FieldOffset(0)]
            public MouseInputData mi;

            [FieldOffset(0)]
            public KEYBDINPUT ki;

            [FieldOffset(0)]
            public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct HARDWAREINPUT
        {
            public int uMsg;
            public short wParamL;
            public short wParamH;
        }

        public struct MouseInputData
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public MouseEventFlags dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [Flags]
        public enum MouseEventFlags : uint
        {
            MOUSEEVENTF_MOVE = 0x0001,
            MOUSEEVENTF_LEFTDOWN = 0x0002,
            MOUSEEVENTF_LEFTUP = 0x0004,
            MOUSEEVENTF_RIGHTDOWN = 0x0008,
            MOUSEEVENTF_RIGHTUP = 0x0010,
            MOUSEEVENTF_MIDDLEDOWN = 0x0020,
            MOUSEEVENTF_MIDDLEUP = 0x0040,
            MOUSEEVENTF_XDOWN = 0x0080,
            MOUSEEVENTF_XUP = 0x0100,
            MOUSEEVENTF_WHEEL = 0x0800,
            MOUSEEVENTF_VIRTUALDESK = 0x4000,
            MOUSEEVENTF_ABSOLUTE = 0x8000
        }

        public enum SendInputEventType : int
        {
            InputMouse,
            InputKeyboard,
            InputHardware
        }

        enum SystemMetric
        {
            SM_CXSCREEN = 0,
            SM_CYSCREEN = 1,
        }

        public enum VirtKeyCodes
        {
            VK_CONTROL = 0x11, 
            VK_SHIFT = 0x10
        }

        [Flags]
        public enum KeyboardEventFlags: uint
        {
            KEYEVENTF_EXTENDEDKEY = 0x0001, 
            KEYEVENTF_KEYUP = 0x0002,
            KEYEVENTF_SCANCODE = 0x0008,
            KEYEVENTF_UNICODE = 0x0004
        }

        [DllImport("user32.dll")]
        static extern int GetSystemMetrics(SystemMetric smIndex);

        static int CalculateAbsoluteCoordinateX(int x)
        {
            return (x * 65536) / GetSystemMetrics(SystemMetric.SM_CXSCREEN);
        }

        static int CalculateAbsoluteCoordinateY(int y)
        {
            return (y * 65536) / GetSystemMetrics(SystemMetric.SM_CYSCREEN);
        }

        private static void KeysDown(int keys)
        {
            if (keys == 0)
            {
                return;
            }

            if (keys == 1) // Control is being pressed
            {
                INPUT keybdInput = new INPUT();
                keybdInput.type = SendInputEventType.InputKeyboard;
                keybdInput.mkhi.ki.wVk = (ushort)VirtKeyCodes.VK_CONTROL;
                keybdInput.mkhi.ki.dwFlags = 0; // key is being pressed
                keybdInput.mkhi.ki.wScan = 0;
                keybdInput.mkhi.ki.time = 0;
                keybdInput.mkhi.ki.dwExtraInfo = IntPtr.Zero;

                SendInput(1, ref keybdInput, Marshal.SizeOf(new INPUT()));
            }
            else if (keys == 2) // Shift is being pressed
            {
                INPUT keybdInput = new INPUT();
                keybdInput.type = SendInputEventType.InputKeyboard;
                keybdInput.mkhi.ki.wVk = (ushort)VirtKeyCodes.VK_SHIFT;
                keybdInput.mkhi.ki.dwFlags = 0; // key is being pressed
                keybdInput.mkhi.ki.wScan = 0;
                keybdInput.mkhi.ki.time = 0;
                keybdInput.mkhi.ki.dwExtraInfo = IntPtr.Zero;

                SendInput(1, ref keybdInput, Marshal.SizeOf(new INPUT()));
            }
            else if (keys == 3) // both Control and Shift keys are being pressed
            {
                INPUT keybdInput = new INPUT();
                keybdInput.type = SendInputEventType.InputKeyboard;
                keybdInput.mkhi.ki.wVk = (ushort)VirtKeyCodes.VK_CONTROL;
                keybdInput.mkhi.ki.dwFlags = 0; // Ctrl key is being pressed
                keybdInput.mkhi.ki.wScan = 0;
                keybdInput.mkhi.ki.time = 0;
                keybdInput.mkhi.ki.dwExtraInfo = IntPtr.Zero;

                SendInput(1, ref keybdInput, Marshal.SizeOf(new INPUT()));

                keybdInput.mkhi.ki.wVk = (ushort)VirtKeyCodes.VK_SHIFT; // press Shift

                SendInput(1, ref keybdInput, Marshal.SizeOf(new INPUT()));
            }
        }

        private static void KeysUp(int keys)
        {
            if (keys == 0)
            {
                return;
            }

            if (keys == 1) // Control is being released
            {
                INPUT keybdInput = new INPUT();
                keybdInput.type = SendInputEventType.InputKeyboard;
                keybdInput.mkhi.ki.wScan = 0;
                keybdInput.mkhi.ki.wVk = (ushort)VirtKeyCodes.VK_CONTROL;
                keybdInput.mkhi.ki.dwFlags = (uint)KeyboardEventFlags.KEYEVENTF_KEYUP;
                keybdInput.mkhi.ki.time = 0;
                keybdInput.mkhi.ki.dwExtraInfo = IntPtr.Zero;

                SendInput(1, ref keybdInput, Marshal.SizeOf(new INPUT()));
            }
            else if (keys == 2) // Shift is being released
            {
                INPUT keybdInput = new INPUT();
                keybdInput.type = SendInputEventType.InputKeyboard;
                keybdInput.mkhi.ki.wScan = 0;
                keybdInput.mkhi.ki.wVk = (ushort)VirtKeyCodes.VK_SHIFT;
                keybdInput.mkhi.ki.dwFlags = (uint)KeyboardEventFlags.KEYEVENTF_KEYUP;
                keybdInput.mkhi.ki.time = 0;
                keybdInput.mkhi.ki.dwExtraInfo = IntPtr.Zero;

                SendInput(1, ref keybdInput, Marshal.SizeOf(new INPUT()));
            }
            else if (keys == 3) // both Control and Shift are being released
            {
                INPUT keybdInput = new INPUT();
                keybdInput.type = SendInputEventType.InputKeyboard;
                keybdInput.mkhi.ki.wScan = 0;
                keybdInput.mkhi.ki.wVk = (ushort)VirtKeyCodes.VK_SHIFT; // Shift up
                keybdInput.mkhi.ki.dwFlags = (uint)KeyboardEventFlags.KEYEVENTF_KEYUP;
                keybdInput.mkhi.ki.time = 0;
                keybdInput.mkhi.ki.dwExtraInfo = IntPtr.Zero;

                SendInput(1, ref keybdInput, Marshal.SizeOf(new INPUT()));

                keybdInput.mkhi.ki.wVk = (ushort)VirtKeyCodes.VK_CONTROL; // Control up

                SendInput(1, ref keybdInput, Marshal.SizeOf(new INPUT()));
            }
        }
		
		public static void MoveMousePointer(int x, int y, int keys)
        {
            INPUT mouseInput = new INPUT();
            mouseInput.type = SendInputEventType.InputMouse;
            mouseInput.mkhi.mi.dx = CalculateAbsoluteCoordinateX(x);
            mouseInput.mkhi.mi.dy = CalculateAbsoluteCoordinateY(y);
            mouseInput.mkhi.mi.mouseData = 0;

            SendInputClass.KeysDown(keys);

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_MOVE | MouseEventFlags.MOUSEEVENTF_ABSOLUTE;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));

            SendInputClass.KeysUp(keys);
        }
		
		public static void MouseScroll(uint wheelTicks)
        {
            INPUT mouseInput = new INPUT();
            mouseInput.type = SendInputEventType.InputMouse;
			mouseInput.mkhi.mi.dx = 0;
            mouseInput.mkhi.mi.dy = 0;
			mouseInput.mkhi.mi.time = 0;
			mouseInput.mkhi.mi.dwExtraInfo = IntPtr.Zero;
            mouseInput.mkhi.mi.mouseData = wheelTicks * 120;
            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_WHEEL;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));
        }

        public static void ClickLeftMouseButton(int x, int y, int keys)
        {
            INPUT mouseInput = new INPUT();
            mouseInput.type = SendInputEventType.InputMouse;
            mouseInput.mkhi.mi.dx = CalculateAbsoluteCoordinateX(x);
            mouseInput.mkhi.mi.dy = CalculateAbsoluteCoordinateY(y);
            mouseInput.mkhi.mi.mouseData = 0;

            SendInputClass.KeysDown(keys);

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_MOVE | MouseEventFlags.MOUSEEVENTF_ABSOLUTE;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_LEFTDOWN;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_LEFTUP;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));

            SendInputClass.KeysUp(keys);
        }

        // x, y - screen coordinates
        // keys = 0 - no key pressed, keys = 1 - Control key pressed,
        // keys = 2 - Shift key pressed, keys = 3 - Control and Shift keys pressed
        public static void ClickRightMouseButton(int x, int y, int keys)
        {
            INPUT mouseInput = new INPUT();
            mouseInput.type = SendInputEventType.InputMouse;
            mouseInput.mkhi.mi.dx = CalculateAbsoluteCoordinateX(x);
            mouseInput.mkhi.mi.dy = CalculateAbsoluteCoordinateY(y);
            mouseInput.mkhi.mi.mouseData = 0;

            SendInputClass.KeysDown(keys);

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_MOVE | MouseEventFlags.MOUSEEVENTF_ABSOLUTE;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_RIGHTDOWN;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_RIGHTUP;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));

            SendInputClass.KeysUp(keys);
        }

        public static void ClickMiddleMouseButton(int x, int y, int keys)
        {
            INPUT mouseInput = new INPUT();
            mouseInput.type = SendInputEventType.InputMouse;
            mouseInput.mkhi.mi.dx = CalculateAbsoluteCoordinateX(x);
            mouseInput.mkhi.mi.dy = CalculateAbsoluteCoordinateY(y);
            mouseInput.mkhi.mi.mouseData = 0;

            SendInputClass.KeysDown(keys);

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_MOVE | MouseEventFlags.MOUSEEVENTF_ABSOLUTE;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_MIDDLEDOWN;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_MIDDLEUP;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));

            SendInputClass.KeysUp(keys);
        }

        public static void DoubleClick(int x, int y, int keys)
        {
            INPUT mouseInput = new INPUT();
            mouseInput.type = SendInputEventType.InputMouse;
            mouseInput.mkhi.mi.dx = CalculateAbsoluteCoordinateX(x);
            mouseInput.mkhi.mi.dy = CalculateAbsoluteCoordinateY(y);
            mouseInput.mkhi.mi.mouseData = 0;

			SendInputClass.KeysDown(keys);

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_MOVE | MouseEventFlags.MOUSEEVENTF_ABSOLUTE;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_LEFTDOWN;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_LEFTUP;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));
          
            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_LEFTDOWN;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));
          
            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_LEFTUP;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));
			
			SendInputClass.KeysUp(keys);
        }
    }

    internal struct WINDOWPLACEMENT
    {
        public int length;
        public int flags;
        public int showCmd;
        public System.Drawing.Point ptMinPosition;
        public System.Drawing.Point ptMaxPosition;
        public System.Drawing.Rectangle rcNormalPosition;
    }

    internal class WindowPlacementConstants
    { 
        public static UInt32 SW_HIDE =         0;
        public static UInt32 SW_SHOWNORMAL =       1;
        public static UInt32 SW_NORMAL =       1;
        public static UInt32 SW_SHOWMINIMIZED =    2;
        public static UInt32 SW_SHOWMAXIMIZED =    3;
        public static UInt32 SW_MAXIMIZE =     3;
        public static UInt32 SW_SHOWNOACTIVATE =   4;
        public static UInt32 SW_SHOW =         5;
        public static UInt32 SW_MINIMIZE =     6;
        public static UInt32 SW_SHOWMINNOACTIVE =  7;
        public static UInt32 SW_SHOWNA =       8;
        public static UInt32 SW_RESTORE =      9;
    }
    
    internal class Win32CalendarMessages
    {
        public static uint MCM_GETCURSEL = 0x1001;
        public static uint MCM_SETCURSEL = 0x1002;
        public static uint MCM_GETSELRANGE = 0x1005;
        public static uint MCM_SETSELRANGE = 0x1006;
    }
    
    internal class DateTimePicker32Messages
    {
        public static uint DTM_GETSYSTEMTIME = 0x1001;
        public static uint DTM_SETSYSTEMTIME = 0x1002;
    }
    
    internal class DateTimePicker32Constants
    {
        public static uint GDT_VALID = 0;
        public static uint GDT_NONE = 1;
    }
    
    [StructLayout(LayoutKind.Sequential)]
    internal struct SYSTEMTIME 
    {
        [MarshalAs(UnmanagedType.U2)] public short Year;
        [MarshalAs(UnmanagedType.U2)] public short Month;
        [MarshalAs(UnmanagedType.U2)] public short DayOfWeek;
        [MarshalAs(UnmanagedType.U2)] public short Day;
        [MarshalAs(UnmanagedType.U2)] public short Hour;
        [MarshalAs(UnmanagedType.U2)] public short Minute;
        [MarshalAs(UnmanagedType.U2)] public short Second;
        [MarshalAs(UnmanagedType.U2)] public short Milliseconds;
    }
    
    internal enum GWL
    {
        GWL_WNDPROC =    (-4),
        GWL_HINSTANCE =  (-6),
        GWL_HWNDPARENT = (-8),
        GWL_STYLE =      (-16),
        GWL_EXSTYLE =    (-20),
        GWL_USERDATA =   (-21),
        GWL_ID =     (-12)
    }
    
    internal class Win32CalendarStyles
    {
        public static uint MCS_MULTISELECT = 2;
    }
}
