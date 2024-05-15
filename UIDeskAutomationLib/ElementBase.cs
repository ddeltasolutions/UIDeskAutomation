using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Base class for all UI controls. This class cannot be instantiated.
    /// </summary>
    abstract public partial class ElementBase
    {
        internal ElementBase() {}

        /// <summary>
        /// inner MS UI Automation element
        /// </summary>
        internal AutomationElement uiElement = null;

        //wait period
        static internal int waitPeriod = 100;
		
		/// <summary>
        /// Gets the inner AutomationElement of this element
        /// </summary>
		public AutomationElement InnerElement
		{
			get
			{
				return this.uiElement;
			}
		}

        /// <summary>
        /// Gets a general description for this element
        /// </summary>
        /// <returns>description for the current element</returns>
        public string GetDescription()
        {
            if (this.uiElement == null)
            {
                Engine.TraceInLogFile("GetDescription method - Null element");
                return "Null element";
            }

            string sName = "";
            try
            {
                sName = this.uiElement.Current.Name;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            ControlType ctrlType = ControlType.Custom;
            try
            {
                ctrlType = this.uiElement.Current.ControlType;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return ctrlType.LocalizedControlType + ": " + sName;
        }

        /// <summary>
        /// Highlights an element for a second
        /// </summary>
        public void Highlight()
        {
            if (this.uiElement == null)
            {
                Engine.TraceInLogFile("Highlight method - Null element");
                throw new Exception("Highlight method - Null element");
            }

            System.Windows.Rect? rect = null;
            try
            {
                rect = this.uiElement.Current.BoundingRectangle;
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("Highlight method threw an exception - " +
                    ex.Message);
                return;
            }

            //if (rect == null)
            if (rect.HasValue == false)
            {
                Engine.TraceInLogFile("Highlight method - cannot obtain BoundingRectangle");
                return;
            }

            int left = 0;
            int top = 0;
            int width = 0;
            int height = 0;

            try
            {
                left = Convert.ToInt32(rect.Value.Left);
                top = Convert.ToInt32(rect.Value.Top);
                width = Convert.ToInt32(rect.Value.Width);
                height = Convert.ToInt32(rect.Value.Height);
            }
            catch
            {
                return;
            }

            int thickness = 5;

            GraphicsPath path = new GraphicsPath();

            path.AddRectangle(new System.Drawing.Rectangle(0, 0, thickness, height)); //add left
            path.AddRectangle(new System.Drawing.Rectangle(thickness, height - thickness, 
                width - 2 * thickness, thickness)); //add bottom
            path.AddRectangle(new System.Drawing.Rectangle(width - thickness, 0, thickness, height)); //add right
            path.AddRectangle(new System.Drawing.Rectangle(thickness, 0, width - 2 * thickness, thickness)); //add top

            System.Drawing.Region region = new System.Drawing.Region(path);
            try
            {
                FrmDummy dummyForm = new FrmDummy();

                dummyForm.Left = left;
                dummyForm.Top = top;

                dummyForm.Width = width;
                dummyForm.Height = height;

                dummyForm.Region = region;

                dummyForm.ShowInTaskbar = false;
                dummyForm.TopMost = true;
                dummyForm.Visible = true;

                Thread.Sleep(200);
                dummyForm.Visible = false;
                Thread.Sleep(200);
                dummyForm.Visible = true;
                Thread.Sleep(200);
                dummyForm.Visible = false;
            }
            catch (Exception ex)
            {
                return;
            }
        }

        /// <summary>
        /// Tests if underlying UI object is still available. 
        /// For example if container window closes, the current UI element is not available to user anymore.
        /// </summary>
        public bool IsAlive
        {
            get
            {
                bool isAlive = true;

                try
                {
                    int processId = this.uiElement.Current.ProcessId;
                }
                catch //ElementNotAvailableException
                {
                    isAlive = false;
                }

                return isAlive;
            }
        }

        /// <summary>
        /// Gets the text of this element.
        /// </summary>
        /// <returns>text of the element</returns>
        public string GetText()
        {
            ValuePattern valuePattern = Helper.GetValuePattern(this.uiElement);

            if (valuePattern != null)
            {
                try
                {
                    return valuePattern.Current.Value;
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("ValuePattern.Value: " + ex.Message);
                }
            }

			string name = null;
            try
            {
                name = this.uiElement.Current.Name;
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("AutomationElement.Name: " + ex.Message);
            }

            if (name != null)
            {
                return name;
            }

            IntPtr windowHandle = this.GetWindow();

            if (windowHandle != IntPtr.Zero)
            {
                UIDA_Window window = new UIDA_Window(windowHandle);
                return window.GetWindowText();
            }
            else
            {
                Engine.TraceInLogFile("GetText error - cannot get element text");
                throw new Exception("GetText error - cannot get element text");
            }
        }

        /// <summary>
        /// Gets the Id of the associated process
        /// </summary>
        public int ProcessId
        {
            get
            {
                IntPtr hWnd = this.GetWindow();
                if (hWnd == IntPtr.Zero)
                {
                    return 0;
                }

                uint processId = 0;
                UnsafeNativeFunctions.GetWindowThreadProcessId(hWnd, out processId);
                return (int)processId;
            }
        }
		
		/// <summary>
        /// Gets the UI Automation Class Name of this element
        /// </summary>
        public string ClassName
        {
            get
            {
                return this.uiElement.Current.ClassName;
            }
        }
		
		/// <summary>
        /// Gets the Left screen coordinate of the element or -1 if the element doesn't have a visual representation.
        /// </summary>
		public int Left
		{
			get
			{
				try
				{
					return (int)this.uiElement.Current.BoundingRectangle.Left;
				}
				catch { }
				
				return -1;
			}
		}
		
		/// <summary>
        /// Gets the Top screen coordinate of the element or -1 if the element doesn't have a visual representation.
        /// </summary>
		public int Top
		{
			get
			{
				try
				{
					return (int)this.uiElement.Current.BoundingRectangle.Top;
				}
				catch { }
				
				return -1;
			}
		}
		
		/// <summary>
        /// Gets the Width of the element bounding rectangle or -1 if the element doesn't have a visual representation.
        /// </summary>
		public int Width
		{
			get
			{
				try
				{
					return (int)this.uiElement.Current.BoundingRectangle.Width;
				}
				catch { }
				
				return -1;
			}
		}
		
		/// <summary>
        /// Gets the Height of the element bounding rectangle or -1 if the element doesn't have a visual representation.
        /// </summary>
		public int Height
		{
			get
			{
				try
				{
					return (int)this.uiElement.Current.BoundingRectangle.Height;
				}
				catch { }
				
				return -1;
			}
		}
		
		/// <summary>
        /// Gets the parent of the element. You need to cast the parent to the proper type. For example, if the current 
		/// element is a list item you can write: "UIDA_List list = listItem.Parent as UIDA_List;"
        /// </summary>
		public ElementBase Parent
		{
			get
			{
				TreeWalker tw = TreeWalker.ControlViewWalker;
				AutomationElement parent = null;
				
				try
				{
					parent = tw.GetParent(this.uiElement);
				}
				catch { }
				
				if (parent == null)
				{
					return null;
				}
				
				ControlType controlType = parent.Current.ControlType;
				if (controlType == ControlType.Button)
				{
					return new UIDA_Button(parent);
				}
				else if (controlType == ControlType.Calendar)
				{
					return new UIDA_Calendar(parent);
				}
				else if (controlType == ControlType.CheckBox)
				{
					return new UIDA_CheckBox(parent);
				}
				else if (controlType == ControlType.ComboBox)
				{
					return new UIDA_ComboBox(parent);
				}
				else if (controlType == ControlType.Hyperlink)
				{
					return new UIDA_HyperLink(parent);
				}
				else if (controlType == ControlType.Image)
				{
					return new UIDA_Image(parent);
				}
				else if (controlType == ControlType.ListItem)
				{
					return new UIDA_ListItem(parent);
				}
				else if (controlType == ControlType.List)
				{
					return new UIDA_List(parent);
				}
				else if (controlType == ControlType.Menu)
				{
					return new UIDA_TopLevelMenu(parent);
				}
				else if (controlType == ControlType.MenuBar)
				{
					return new UIDA_MenuBar(parent);
				}
				else if (controlType == ControlType.MenuItem)
				{
					return new UIDA_MenuItem(parent);
				}
				else if (controlType == ControlType.ProgressBar)
				{
					return new UIDA_ProgressBar(parent);
				}
				else if (controlType == ControlType.RadioButton)
				{
					return new UIDA_RadioButton(parent);
				}
				else if (controlType == ControlType.ScrollBar)
				{
					return new UIDA_ScrollBar(parent);
				}
				else if (controlType == ControlType.Slider)
				{
					return new UIDA_Slider(parent);
				}
				else if (controlType == ControlType.Spinner)
				{
					return new UIDA_Spinner(parent);
				}
				else if (controlType == ControlType.StatusBar)
				{
					return new UIDA_StatusBar(parent);
				}
				else if (controlType == ControlType.Tab)
				{
					return new UIDA_TabCtrl(parent);
				}
				else if (controlType == ControlType.TabItem)
				{
					return new UIDA_TabItem(parent);
				}
				else if (controlType == ControlType.Text)
				{
					return new UIDA_Label(parent);
				}
				else if (controlType == ControlType.ToolBar)
				{
					return new UIDA_Toolbar(parent);
				}
				else if (controlType == ControlType.ToolTip)
				{
					return new UIDA_Tooltip(parent);
				}
				else if (controlType == ControlType.Tree)
				{
					return new UIDA_Tree(parent);
				}
				else if (controlType == ControlType.TreeItem)
				{
					return new UIDA_TreeItem(parent);
				}
				else if (controlType == ControlType.Custom)
				{
					return new UIDA_Custom(parent);
				}
				else if (controlType == ControlType.Group)
				{
					return new UIDA_Group(parent);
				}
				else if (controlType == ControlType.Thumb)
				{
					return new UIDA_Thumb(parent);
				}
				else if (controlType == ControlType.DataGrid)
				{
					return new UIDA_DataGrid(parent);
				}
				else if (controlType == ControlType.DataItem)
				{
					return new UIDA_DataItem(parent);
				}
				else if (controlType == ControlType.Document)
				{
					return new UIDA_Document(parent);
				}
				else if (controlType == ControlType.SplitButton)
				{
					return new UIDA_SplitButton(parent);
				}
				else if (controlType == ControlType.Window)
				{
					return new UIDA_Window(parent);
				}
				else if (controlType == ControlType.Pane)
				{
					return new UIDA_Pane(parent);
				}
				else if (controlType == ControlType.Header)
				{
					return new UIDA_Header(parent);
				}
				else if (controlType == ControlType.HeaderItem)
				{
					return new UIDA_HeaderItem(parent);
				}
				else if (controlType == ControlType.Table)
				{
					return new UIDA_Table(parent);
				}
				else if (controlType == ControlType.TitleBar)
				{
					return new UIDA_TitleBar(parent);
				}
				else if (controlType == ControlType.Separator)
				{
					return new UIDA_Separator(parent);
				}
				
				return null;
			}
		}
		
		/// <summary>
        /// Gets a boolean to determine if the UI element is enabled.
        /// </summary>
		public bool IsEnabled
		{
			get
			{
				try
				{
					return this.uiElement.Current.IsEnabled;
				}
				catch 
				{
					return true;
				}
			}
		}
    }
}
