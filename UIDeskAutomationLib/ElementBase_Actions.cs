using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using System.Threading;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows;

namespace UIDeskAutomationLib
{
    partial class ElementBase
    {
        #region Actions
        
        /// <summary>
        /// Invoke the default action on element
        /// </summary>
        public void Invoke()
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("Invoke - Element not available anymore.");
                throw new Exception("Invoke - Element not available anymore.");
            }

            object objectPattern = null;
            this.uiElement.TryGetCurrentPattern(InvokePattern.Pattern, out objectPattern);
            InvokePattern invokePattern = objectPattern as InvokePattern;

            if (invokePattern == null)
            {
                //Engine.TraceInLogFile("Invoke method - Invoke pattern not supported");
                //throw new Exception("Invoke method - Invoke pattern not supported");

                Engine.TraceInLogFile("Invoke method - Invoke pattern not supported, Try TogglePattern");
				
                object objTogglePattern = null;
                this.uiElement.TryGetCurrentPattern(TogglePattern.Pattern, out objTogglePattern);
                TogglePattern togglePattern = objTogglePattern as TogglePattern;
                if (togglePattern != null)
                {
                    togglePattern.Toggle();
                }
                else
                {
                    Engine.TraceInLogFile("Invoke method - TogglePattern not supported");
                }
                return;
            }

            try
            {
                invokePattern.Invoke();
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Invoke method - " + ex.Message);
                throw new Exception("Invoke method - " + ex.Message);
            }
        }

        /// <summary>
        /// Sets focus to current element
        /// </summary>
        public void Focus()
        {
            try
            {
                this.uiElement.SetFocus();
            }
            catch (Exception ex)
            {
                //IntPtr hwnd = new IntPtr(this.uiElement.Current.NativeWindowHandle);
                IntPtr hwnd = this.GetWindow();

                if (hwnd == IntPtr.Zero)
                {
                    Engine.TraceInLogFile("Focus failed: Cannot get element window.");
                    throw new Exception("Focus failed: Cannot get element window.");
                }

                // first try to bring the window in foreground
                UnsafeNativeFunctions.SetForegroundWindow(hwnd);

                uint processId = 0;
                uint windowThreadId = UnsafeNativeFunctions.GetWindowThreadProcessId(
                    hwnd, out processId);

                uint currentThreadId = UnsafeNativeFunctions.GetCurrentThreadId();

                bool bResult = UnsafeNativeFunctions.AttachThreadInput(
                    currentThreadId, windowThreadId, true);
                Debug.Assert(bResult == true);

                IntPtr hwndResult = UnsafeNativeFunctions.SetFocus(hwnd);

                bResult = UnsafeNativeFunctions.AttachThreadInput(currentThreadId,
                    windowThreadId, false);
                Debug.Assert(bResult == true);

                //if (hwndResult == IntPtr.Zero)
                //{
                //    Engine.TraceInLogFile("Focus failed: " + ex.Message);
                //    throw ex;
                //}
            }
        }

        /// <summary>
        /// Sends keystrokes to this element. Main window must be in foreground. 
		/// This function behaves like the .NET 'SendKeys.Send(string)' function.
		/// More details can be found here: https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys.send?view=netframework-4.8
        /// </summary>
        /// <param name="text">keys to send</param>
        public void SendKeys(string text)
        {
            try
            {
                this.Focus();
                System.Windows.Forms.SendKeys.SendWait(text);
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("SendKeys failed: " + ex.Message);
                throw ex;
            }
        }
		
		/// <summary>
        /// Simulates sending keys to this element using Windows messages.
		/// This function behaves like the .NET 'SendKeys.Send(string)' function.
        /// </summary>
        /// <param name="text">keys to simulate</param>
        public void SimulateSendKeys(string text)
        {
            IntPtr hwnd = this.GetWindow();

            if (hwnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("Cannot get window handle of current element.");
                throw new Exception("Cannot get window handle of current element.");
            }

            Helper.SimulateSendKeys(text, hwnd);
        }

        internal IntPtr GetWindow()
        {
            TreeWalker tw = TreeWalker.ControlViewWalker;

            int hwndInt = 0;
            AutomationElement currentElement = this.uiElement;

            while (true)
            {
                try
                {
                    hwndInt = currentElement.Current.NativeWindowHandle;
                }
                catch (ElementNotAvailableException ex)
                {
                    break;
                }

                if (hwndInt != 0)
                {
                    break;
                }

                currentElement = tw.GetParent(currentElement);
                if (currentElement == null)
                {
                    break;
                }
            }

            return new IntPtr(hwndInt);
        }

        /// <summary>
        /// Left mouse click the element at a clickable position.
        /// </summary>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        public void Click(int keys = 0)
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("This UI element is not available anymore.");
                throw new Exception("This UI element is not available anymore.");
            }

            this.Click(1, keys);
        }

        /// <summary>
        /// Right mouse click the element at a clickable position.
        /// </summary>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        public void RightClick(int keys = 0)
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("RightClick - This UI element is not available anymore.");
                throw new Exception("RightClick - This UI element is not available anymore.");
            }

            this.Click(2, keys);
        }

        /// <summary>
        /// Clicks the element with the middle mouse button at a clickable position.
        /// </summary>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        public void MiddleClick(int keys = 0)
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("MiddleClick - This UI element is not available anymore.");
                throw new Exception("MiddleClick - This UI element is not available anymore.");
            }

            this.Click(3, keys);
        }

        /// <summary>
        /// Mouse click an element.
        /// </summary>
        /// <param name="mouseButton">mouse button, 1 - left button, 2 - right button, 3 - middle button</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        private void Click(int mouseButton, int keys)
        {
            System.Drawing.Point? clickablePoint =
                GetElementClickablePointScreenCoordinates();

            if (clickablePoint.HasValue == true)
            {
                if (Engine.GetInstance() != null)
                {
                    if (mouseButton == 1) // left click
                    {
                        Engine.GetInstance().ClickScreenCoordinatesAt(
                            clickablePoint.Value.X, clickablePoint.Value.Y, keys);
                    }
                    else if (mouseButton == 2) // right click
                    {
                        Engine.GetInstance().RightClickScreenCoordinatesAt(
                            clickablePoint.Value.X, clickablePoint.Value.Y, keys);
                    }
                    else if (mouseButton == 3) // middle click
                    {
                        Engine.GetInstance().MiddleClickScreenCoordinatesAt(
                            clickablePoint.Value.X, clickablePoint.Value.Y, keys);
                    }
                }
            }
            else
            {
                Engine.TraceInLogFile("Cannot get a clickable point.");
                throw new Exception("Cannot get a clickable point.");
            }
        }

        private System.Drawing.Point? GetElementClickablePointScreenCoordinates()
        {
            int x = -1;
            int y = -1;

            System.Windows.Point point = new System.Windows.Point();
            if (this.uiElement.TryGetClickablePoint(out point) == true)
            {
                x = (int)point.X;
                y = (int)point.Y;
            }
            else
            {
                System.Windows.Rect? boundingRectangle = null;

                try
                {
                    boundingRectangle = this.uiElement.Current.BoundingRectangle;
                }
                catch { }

                if (boundingRectangle.HasValue)
                {
                    x = (int)((boundingRectangle.Value.Left + boundingRectangle.Value.Right) / 2.0);
                    y = (int)((boundingRectangle.Value.Top + boundingRectangle.Value.Bottom) / 2.0);
                }
            }

            if (x >= 0)
            {
                return new System.Drawing.Point(x, y);
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Double left mouse button click the element at a clickable position.
        /// </summary>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        public void DoubleClick(int keys = 0)
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("DoubleClick - This UI element is not available anymore.");
                throw new Exception("DoubleClick - This UI element is not available anymore.");
            }

            this.Click(keys);
            this.Click(keys);
        }

        /// <summary>
        /// Left click the element at a specified position.
        /// </summary>
        /// <param name="x">x coordinate relative to this element</param>
        /// <param name="y">y coordinate relative to this element</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift pressed</param>
        public void ClickAt(int x, int y, int keys = 0)
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("ClickAt - This UI element is not available anymore.");
                throw new Exception("ClickAt - This UI element is not available anymore.");
            }

            this.ClickAt(x, y, 1, keys);
        }

        /// <summary>
        /// Right click the element at a specified position.
        /// </summary>
        /// <param name="x">x coordinate relative to this element</param>
        /// <param name="y">y coordinate relative to this element</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        public void RightClickAt(int x, int y, int keys = 0)
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("RightClickAt - This UI element is not available anymore.");
                throw new Exception("RightClickAt - This UI element is not available anymore.");
            }

            this.ClickAt(x, y, 2, keys);
        }

        /// <summary>
        /// Clicks the element with the middle mouse button at a specified position.
        /// </summary>
        /// <param name="x">x coordinate relative to this element</param>
        /// <param name="y">y coordinate relative to this element</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift pressed</param>
        public void MiddleClickAt(int x, int y, int keys = 0)
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("MiddleClickAt - This UI element is not available anymore.");
                throw new Exception("MiddleClickAt - This UI element is not available anymore.");
            }

            this.ClickAt(x, y, 3, keys);
        }

        /// <summary>
        /// Clicks an element at a specified point.
        /// </summary>
        /// <param name="x">x coordinate relative to this element</param>
        /// <param name="y">y coordinate relative to this element</param>
        /// <param name="mouseButton">mouse button, 1 - left button, 2 - right button, 3 - middle button</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        private void ClickAt(int x, int y, int mouseButton, int keys)
        {
            // transform relative coordinates to screen coordinates

            //get bounding rectangle
            System.Windows.Rect? boundingRect = null;

            try
            {
                boundingRect = this.uiElement.Current.BoundingRectangle;
            }
            catch { }

            if (boundingRect.HasValue == false)
            {
                Engine.TraceInLogFile("Cannot get element coordinates.");
                throw new Exception("Cannot get element coordinates.");
            }

            int xScreen = (int)boundingRect.Value.Left + x;
            int yScreen = (int)boundingRect.Value.Top + y;

            // click at specified screen coordinates
            if (Engine.GetInstance() != null)
            {
                if (mouseButton == 1)
                {
                    Engine.GetInstance().ClickScreenCoordinatesAt(
                        xScreen, yScreen, keys);
                }
                else if (mouseButton == 2)
                {
                    Engine.GetInstance().RightClickScreenCoordinatesAt(
                        xScreen, yScreen, keys);
                }
                else if (mouseButton == 3)
                {
                    Engine.GetInstance().MiddleClickScreenCoordinatesAt(
                        xScreen, yScreen, keys);
                }
            }
        }

        /// <summary>
        /// Double clicks the element at a specified position.
        /// </summary>
        /// <param name="x">x coordinate relative to this element</param>
        /// <param name="y">y coordinate relative to this element</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift pressed</param>
        public void DoubleClickAt(int x, int y, int keys = 0)
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("DoubleClickAt - This UI element is not available anymore.");
                throw new Exception("DoubleClickAt - This UI element is not available anymore.");
            }

            this.ClickAt(x, y, keys);
            this.ClickAt(x, y, keys);
        }
		
		/// <summary>
        /// Moves the mouse pointer at a specified point.
        /// </summary>
        /// <param name="x">x coordinate relative to this element</param>
        /// <param name="y">y coordinate relative to this element</param>
        /// <param name="keys">keys pressed, 0 - None, 1 - Control pressed, 
        /// 2 - Shift pressed, 3 - Both Control and Shift are pressed</param>
        public void MoveMouse(int x, int y, int keys = 0)
		{
			// transform relative coordinates to screen coordinates

            //get bounding rectangle
            System.Windows.Rect? boundingRect = null;

            try
            {
                boundingRect = this.uiElement.Current.BoundingRectangle;
            }
            catch { }

            if (boundingRect.HasValue == false)
            {
                Engine.TraceInLogFile("Cannot get element coordinates.");
                throw new Exception("Cannot get element coordinates.");
            }

            int xScreen = (int)boundingRect.Value.Left + x;
            int yScreen = (int)boundingRect.Value.Top + y;
			
			if (Engine.GetInstance() != null)
			{
				Engine.GetInstance().MoveMouse(xScreen, yScreen, keys);
			}
		}
		
		/// <summary>
        /// Simulates a left mouse button click using Windows messages.
        /// </summary>
        /// <param name="keys">keys pressed, 0 - None keys pressed, 1 - Ctrl key pressed, 2 - Shift key pressed, 3 - both Ctrl and Shift keys pressed</param>
        public void SimulateClick(int keys = 0)
        {
            this.SimulateClick(1, keys);
        }

        /// <summary>
        /// Simulates a right mouse button click using Windows messages.
        /// </summary>
        /// <param name="keys">keys pressed, 0 - None keys pressed, 1 - Ctrl key pressed, 2 - Shift key pressed, 3 - both Ctrl and Shift keys pressed</param>
        public void SimulateRightClick(int keys = 0)
        {
            this.SimulateClick(2, keys);
        }

        /// <summary>
        /// Simulates a middle mouse button click using Windows messages.
        /// </summary>
        /// <param name="keys">keys pressed, 0 - None keys pressed, 1 - Ctrl key pressed, 2 - Shift key pressed, 3 - both Ctrl and Shift keys pressed</param>
        public void SimulateMiddleClick(int keys = 0)
        {
            this.SimulateClick(3, keys);
        }

        /// <summary>
        /// Simulates a left mouse button double click using Windows messages.
        /// </summary>
        /// <param name="keys">keys pressed, 0 - None keys pressed, 1 - Ctrl key pressed, 2 - Shift key pressed, 3 - both Ctrl and Shift keys pressed</param>
        public void SimulateDoubleClick(int keys = 0)
        {
            this.SimulateClick(4, keys);
        }

        /// <summary>
        /// Simulates a mouse button click using Windows messages.
        /// </summary>
        /// <param name="keys">keys pressed, 0 - None keys pressed, 1 - Ctrl key pressed, 2 - Shift key pressed, 3 - both Ctrl and Shift keys pressed</param>
        /// <param name="mouseButton">1 - left button, 2 - right button, 
        /// 3 - middle button, 4 - double click</param>
        private void SimulateClick(int mouseButton, int keys = 0)
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("SimulateClick - This UI element is not available anymore.");
                throw new Exception("SimulateClick - This UI element is not available anymore.");
            }

            IntPtr hwnd = this.GetWindow();

            if (hwnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("SimulateClick - Cannot simulate click on this element.");
                throw new Exception("SimulateClick - Cannot simulate click on this element.");
            }

            System.Drawing.Point? clickablePointScreen =
                GetElementClickablePointScreenCoordinates();

            if (clickablePointScreen.HasValue == false)
            {
                Engine.TraceInLogFile("SimulateClick - Cannot get clickable point.");
                throw new Exception("SimulateClick - Cannot get clickable point.");
            }

            Helper.SimulateClickAt(mouseButton, clickablePointScreen.Value.X,
                clickablePointScreen.Value.Y, keys, hwnd);
        }

        /// <summary>
        /// Simulates a left mouse button click using Windows messages at specified location.
        /// </summary>
        /// <param name="x">x client coordinate</param>
        /// <param name="y">y client coordinate</param>
        /// <param name="keys">keys pressed, 0 - None keys pressed, 1 - Ctrl key pressed, 2 - Shift key pressed, 3 - both Ctrl and Shift keys pressed</param>
        public void SimulateClickAt(int x, int y, int keys = 0)
        {
            this.SimulateClickAt(x, y, 1, keys);
        }

        /// <summary>
        /// Simulates a right mouse button click using Windows messages at specified location.
        /// </summary>
        /// <param name="x">x client coordinate</param>
        /// <param name="y">y client coordinate</param>
        /// <param name="keys">keys pressed, 0 - None keys pressed, 1 - Ctrl key pressed, 2 - Shift key pressed, 3 - both Ctrl and Shift keys pressed</param>
        public void SimulateRightClickAt(int x, int y, int keys = 0)
        {
            this.SimulateClickAt(x, y, 2, keys);
        }

        /// <summary>
        /// Simulates a middle mouse button click using Windows messages at specified location.
        /// </summary>
        /// <param name="x">x client coordinate</param>
        /// <param name="y">y client coordinate</param>
        /// <param name="keys">keys pressed, 0 - None keys pressed, 1 - Ctrl key pressed, 2 - Shift key pressed, 3 - both Ctrl and Shift keys pressed</param>
        public void SimulateMiddleClickAt(int x, int y, int keys = 0)
        {
            this.SimulateClickAt(x, y, 3, keys);
        }

        /// <summary>
        /// Simulates a left mouse button double click using Windows messages at specified location.
        /// </summary>
        /// <param name="x">x client coordinate</param>
        /// <param name="y">y client coordinate</param>
        /// <param name="keys">keys pressed, 0 - None keys pressed, 1 - Ctrl key pressed, 2 - Shift key pressed, 3 - both Ctrl and Shift keys pressed</param>
        public void SimulateDoubleClickAt(int x, int y, int keys = 0)
        {
            this.SimulateClickAt(x, y, 4, keys);
        }

        /// <summary>
        /// Simulates a mouse button click using Windows messages at specified location.
        /// </summary>
        /// <param name="x">x client coordinate</param>
        /// <param name="y">y client coordinate</param>
        /// <param name="keys">keys pressed, 0 - None keys pressed, 1 - Ctrl key pressed, 2 - Shift key pressed, 3 - both Ctrl and Shift keys pressed</param>
        /// <param name="mouseButton">1 - left button, 2 - right button, 3 - middle button, 4 - double click</param>
        private void SimulateClickAt(int x, int y, int mouseButton, int keys)
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("SimulateClickAt - This element is not available anymore");
                throw new Exception("SimulateClickAt - This element is not available anymore");
            }

            IntPtr hwnd = this.GetWindow();

            if (hwnd == IntPtr.Zero)
            {
                Engine.TraceInLogFile("SimulateClickAt - Cannot get element window");
                throw new Exception("SimulateClickAt - Cannot get element window");
            }

            // transform relative coordinates to screen coordinates

            //get bounding rectangle
            System.Windows.Rect? boundingRect = null;

            try
            {
                boundingRect = this.uiElement.Current.BoundingRectangle;
            }
            catch { }

            if (boundingRect.HasValue == false)
            {
                Engine.TraceInLogFile("SimulateClickAt - Cannot get element coordinates.");
                throw new Exception("SimulateClickAt - Cannot get element coordinates.");
            }

            int xScreen = (int)boundingRect.Value.Left + x;
            int yScreen = (int)boundingRect.Value.Top + y;

            Helper.SimulateClickAt(mouseButton, xScreen, yScreen, keys, hwnd);
        }
		
		/// <summary>
        /// Scrolls mouse wheel up over this element with the specified number of wheel ticks
        /// </summary>
        /// <param name="wheelTicks">number of wheel ticks</param>
        public void MouseScrollUp(uint wheelTicks)
        {
			try
            {
                this.Focus();
				
				Rect boundingRect = this.uiElement.Current.BoundingRectangle;
				int xMiddle = (int)(boundingRect.Left + boundingRect.Right) / 2;
				int yMiddle = (int)(boundingRect.Top + boundingRect.Bottom) / 2;
				
				Engine engine = Engine.GetInstance();
				if (engine != null)
				{
					engine.MoveMouse(xMiddle, yMiddle);
					Thread.Sleep(100);
					engine.MouseScrollUp(wheelTicks);
				}
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("MouseScrollUp failed: " + ex.Message);
                throw ex;
            }
        }
		
		/// <summary>
        /// Scrolls mouse wheel down over this element with the specified number of wheel ticks
        /// </summary>
        /// <param name="wheelTicks">number of wheel ticks</param>
        public void MouseScrollDown(uint wheelTicks)
        {
			try
            {
                this.Focus();
				
				Rect boundingRect = this.uiElement.Current.BoundingRectangle;
				int xMiddle = (int)(boundingRect.Left + boundingRect.Right) / 2;
				int yMiddle = (int)(boundingRect.Top + boundingRect.Bottom) / 2;
				
				Engine engine = Engine.GetInstance();
				if (engine != null)
				{
					engine.MoveMouse(xMiddle, yMiddle);
					Thread.Sleep(100);
					engine.MouseScrollDown(wheelTicks);
				}
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("MouseScrollDown failed: " + ex.Message);
                throw ex;
            }
        }

        /// <summary>
        /// Captures the element from screen and saves the image in the specified file. Supported formats: BMP, JPG/JPEG and PNG. 
		/// Use it when the element may be partially or fully hidden or overlapped by another window.
        /// </summary>
        /// <param name="fileName">Image file path. If only the file name is provided, without a full path, then the image will be saved in the current directory of the application which loaded this library.</param>
		/// <param name="cropRect">Coordinates of the rectangle to crop relatively to this element. Don't specify it if you want to capture the whole element.</param>
        public void CaptureToFile(string fileName, Rect? cropRect = null)
        {
            Bitmap bitmap = Helper.CaptureElement(uiElement);
			
			if (cropRect.HasValue)
			{
				Rectangle cropRectangle = new Rectangle((int)cropRect.Value.X, (int)cropRect.Value.Y, 
					(int)cropRect.Value.Width, (int)cropRect.Value.Height);
				bitmap = Helper.CropImage(bitmap, cropRectangle);
			}
            
            if (bitmap != null)
            {
                ImageFormat format = ImageFormat.Bmp;
                string extension = System.IO.Path.GetExtension(fileName).ToLower();
                if (extension == "jpg" || extension == "jpeg")
                {
                    format = ImageFormat.Jpeg;
                }
                else if (extension == "png")
                {
                    format = ImageFormat.Png;
                }
                
                bitmap.Save(fileName, format);
                bitmap.Dispose();
            }
        }
		
		/// <summary>
        /// Captures the element from screen and saves the image in the specified file. Supported formats: BMP, JPG/JPEG and PNG. 
		/// Use it when the element is entirely visible.
        /// </summary>
        /// <param name="fileName">Image file path. If only the file name is provided, without a full path, then the image will be saved in the current directory of the application which loaded this library.</param>
		/// <param name="cropRect">Coordinates of the rectangle to crop relatively to this element. Don't specify it if you want to capture the whole element.</param>
        public void CaptureVisibleToFile(string fileName, Rect? cropRect = null)
        {
            Bitmap bitmap = Helper.CaptureVisibleElement(uiElement, cropRect);
            
            if (bitmap != null)
            {
                ImageFormat format = ImageFormat.Bmp;
                string extension = System.IO.Path.GetExtension(fileName).ToLower();
                if (extension == "jpg" || extension == "jpeg")
                {
                    format = ImageFormat.Jpeg;
                }
                else if (extension == "png")
                {
                    format = ImageFormat.Png;
                }
                
                bitmap.Save(fileName, format);
                bitmap.Dispose();
            }
        }
		
		/// <summary>
        /// Captures the element from screen and saves the image in an System.Drawing.Bitmap object. 
		/// Use it when the element may be partially or fully hidden or overlapped by another window.
        /// </summary>
		/// <param name="cropRect">Coordinates of the rectangle to crop relatively to this element. Don't specify it if you want to capture the whole element.</param>
        public Bitmap CaptureToBitmap(Rect? cropRect = null)
        {
            Bitmap bitmap = Helper.CaptureElement(uiElement);
			
			if (cropRect.HasValue)
			{
				Rectangle cropRectangle = new Rectangle((int)cropRect.Value.X, (int)cropRect.Value.Y, 
					(int)cropRect.Value.Width, (int)cropRect.Value.Height);
				bitmap = Helper.CropImage(bitmap, cropRectangle);
			}
			
			return bitmap;
		}
		
		/// <summary>
        /// Captures the element from screen and saves the image in an System.Drawing.Bitmap object. 
		/// Use it when the element is entirely visible.
        /// </summary>
		/// <param name="cropRect">Coordinates of the rectangle to crop relatively to this element. Don't specify it if you want to capture the whole element.</param>
        public Bitmap CaptureVisibleToBitmap(Rect? cropRect = null)
        {
            return Helper.CaptureVisibleElement(uiElement, cropRect);
		}
		
		/// <summary>
        /// Waits for the process that created this element to enter an idle state. Don't specify a parameter value for an infinite wait.
        /// </summary>
		/// <param name="milliseconds">Specifies the amount of time, in milliseconds, to wait for the process to become idle. If no value is specified then the function will wait indefinitely.</param>
        public void WaitForInputIdle(int milliseconds = -1)
        {
			Process proc = null;
			try
			{
				proc = Process.GetProcessById(this.ProcessId);
			}
			catch { }
			
			if (proc == null)
			{
				return;
			}
			
			try
			{
				if (milliseconds == -1)
				{
					proc.WaitForInputIdle();
				}
				else
				{
					proc.WaitForInputIdle(milliseconds);
				}
			}
			catch { }
		}

        #endregion
    }
}
