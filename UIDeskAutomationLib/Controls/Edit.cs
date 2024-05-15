using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;
using System.Windows.Automation.Text;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents an Edit UI element
    /// </summary>
    public class UIDA_Edit : ElementBase
    {
        internal UIDA_Edit() {}

        /// <summary>
        /// Creates a UIDA_Edit using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_Edit(AutomationElement el)
        {
            base.uiElement = el;
        }

        /// <summary>
        /// Sets text to this Edit element
        /// </summary>
        /// <param name="text">text to set</param>
        public void SetText(string text)
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("This UI element is not available anymore.");
                throw new Exception("This UI element is not available anymore.");
            }

            object valuePatternObj = null;
            if (this.uiElement.TryGetCurrentPattern(ValuePatternIdentifiers.Pattern, out valuePatternObj) == true)
            {
                ValuePattern valuePattern = valuePatternObj as ValuePattern;
                if (valuePattern != null)
                {
                    if (valuePattern.Current.IsReadOnly == true)
                    {
                        Engine.TraceInLogFile("Edit control is read-only.");
                        throw new Exception("Edit control is read-only");
                    }
                    else
                    {
                        valuePattern.SetValue(text);
                    }
                    return; // text successfully set
                }
            }

            // try with native Win32 function SetWindowText
            IntPtr hwnd = IntPtr.Zero;
            try
            {
                hwnd = new IntPtr(base.uiElement.Current.NativeWindowHandle);
            }
            catch { }

            if (hwnd != IntPtr.Zero)
            {
                IntPtr textPtr = IntPtr.Zero;

                try
                {
                    textPtr = Marshal.StringToBSTR(text);
                }
                catch { }

                try
                {
                    if (textPtr != IntPtr.Zero)
                    {
                        if (UnsafeNativeFunctions.SendMessage(hwnd,
                            WindowMessages.WM_SETTEXT, IntPtr.Zero, textPtr) == Win32Constants.TRUE)
                        {
                            return; // text successfully set
                        }
                    }
                }
                catch { }
                finally
                {
                    Marshal.FreeBSTR(textPtr);
                }
            }

            //simulate send chars
            foreach (char ch in text)
            {
                // send WM_CHAR for each character
                UnsafeNativeFunctions.PostMessage(hwnd, WindowMessages.WM_CHAR,
                    new IntPtr(ch), new IntPtr(1));
            }
        }

        /// <summary>
        /// Gets the text of this Edit element
        /// </summary>
        /// <returns>the text of this Edit element</returns>
        public new string GetText()
        {
            if (this.IsAlive == false)
            {
                Engine.TraceInLogFile("This UI element is not available to the user anymore.");
                throw new Exception("This UI element is not available to the user anymore.");
            }

            object valuePatternObj = null;
            if (this.uiElement.TryGetCurrentPattern(ValuePatternIdentifiers.Pattern, out valuePatternObj) == true)
            {
                ValuePattern valuePattern = valuePatternObj as ValuePattern;
                if (valuePattern != null)
                {
                    return valuePattern.Current.Value;
                }
            }

            object textPatternObject = null;
            if (this.uiElement.TryGetCurrentPattern(TextPatternIdentifiers.Pattern, out textPatternObject) == true)
            {
                TextPattern textPattern = textPatternObject as TextPattern;
                if (textPattern != null)
                {
                    return textPattern.DocumentRange.GetText(-1);
                }
            }

            // try with native Win32 function SetWindowText
            IntPtr hwnd = IntPtr.Zero;
            try
            {
                hwnd = new IntPtr(base.uiElement.Current.NativeWindowHandle);
            }
            catch { }

            if (hwnd != IntPtr.Zero)
            {
                IntPtr textLengthPtr = UnsafeNativeFunctions.SendMessage(hwnd,
                    WindowMessages.WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero);

                if (textLengthPtr.ToInt32() > 0)
                {
                    int textLength = textLengthPtr.ToInt32() + 1;

                    StringBuilder text = new StringBuilder(textLength);
                    UnsafeNativeFunctions.SendMessage(hwnd, WindowMessages.WM_GETTEXT, textLength, text);

                    return text.ToString();
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Gets or sets the control's text
        /// </summary>
        public string Text
        {
            get
            {
                return this.GetText();
            }
            set
            {
                this.SetText(value);
            }
        }

        /// <summary>
        /// Clears the text of an edit control.
        /// </summary>
        public void ClearText()
        {
            this.SetText("");
        }

        /// <summary>
        /// Selects a text in an edit control.
        /// </summary>
        /// <param name="text">text to select</param>
        /// <param name="backwards">true if search backwards in edit text, default false</param>
        /// <param name="ignoreCase">true if case is ignored in search, default true</param>
        public void SelectText(string text, bool backwards = false, bool ignoreCase = true)
        {
            object textPatternObj = null;
            if (this.uiElement.TryGetCurrentPattern(TextPattern.Pattern, out textPatternObj) == true)
            {
                TextPattern textPattern = textPatternObj as TextPattern;
                if (textPattern == null)
                {
                    return;
                }

                if (textPattern.SupportedTextSelection == SupportedTextSelection.None)
                {
                    Engine.TraceInLogFile("SelectText method: selection not supported");
                    throw new Exception("SelectText method: selection not supported");
                }

                TextPatternRange document = null;
                try
                {
                    document = textPattern.DocumentRange;
                }
                catch { }

                if (document == null)
                {
                    return;
                }

                TextPatternRange textRange = document.FindText(text, backwards, ignoreCase);
                if (textRange == null)
                {
                    Engine.TraceInLogFile("SelectText method: Cannot find text");
                    return;
                }

                try
                {
                    textRange.Select();
                }
                catch (Exception ex)
                {
                    Engine.TraceInLogFile("SelectText method: " + ex.Message);
                    throw new Exception("SelectText method: " + ex.Message);
                }
            }
        }
		
		/// <summary>
        /// Selects all text in an edit control.
        /// </summary>
        public void SelectAll()
        {
            object textPatternObj = null;
            this.uiElement.TryGetCurrentPattern(TextPattern.Pattern, out textPatternObj);
			
			TextPattern textPattern = textPatternObj as TextPattern;
			if (textPattern == null)
			{
				// try Ctrl+A
				this.KeyDown(VirtualKeys.Control);
				this.KeyPress(VirtualKeys.A);
				this.KeyUp(VirtualKeys.Control);
				return;
			}

			if (textPattern.SupportedTextSelection == SupportedTextSelection.None)
			{
				Engine.TraceInLogFile("SelectAll method: selection not supported");
				throw new Exception("SelectAll method: selection not supported");
			}

			TextPatternRange document = null;
			try
			{
				document = textPattern.DocumentRange;
			}
			catch { }

			if (document == null)
			{
				return;
			}
			
			try
			{
				document.Select();
			}
			catch (Exception ex)
			{
				Engine.TraceInLogFile("SelectAll method: " + ex.Message);
				throw new Exception("SelectAll method: " + ex.Message);
			}
		}

        /// <summary>
        /// Clears any selected text in an edit control.
        /// </summary>
        public void ClearSelection()
        {
            object textPatternObj = null;
            if (this.uiElement.TryGetCurrentPattern(TextPattern.Pattern, out textPatternObj) == true)
            {
                TextPattern textPattern = textPatternObj as TextPattern;
                if (textPattern == null)
                {
                    return;
                }

                TextPatternRange[] selections = null;
                try
                {
                    selections = textPattern.GetSelection();
                }
                catch { }

                if (selections == null)
                {
                    return;
                }

                foreach (TextPatternRange selection in selections)
                {
                    try
                    {
                        selection.RemoveFromSelection();
                    }
                    catch (Exception ex)
                    {
                        Engine.TraceInLogFile("ClearSelection method: operation not supported");
                        throw new Exception("ClearSelection method: operation not supported");
                    }
                }
            }
        }
		
		/// <summary>
        /// Gets the selected text from the edit control. Returns null if cannot get the selected text.
        /// </summary>
		/// <returns>the selected text of this Edit element</returns>
		public string GetSelectedText()
		{
			try
			{
				object objTextPattern = null;
				if (this.uiElement.TryGetCurrentPattern(TextPatternIdentifiers.Pattern, out objTextPattern) == true)
				{
					TextPattern textPattern = objTextPattern as TextPattern;
					if (textPattern != null)
					{
						TextPatternRange[] selectionRanges = textPattern.GetSelection();
						if (selectionRanges.Length > 0)
						{
							return selectionRanges[0].GetText(-1);
						}
					}
				}
			}
			catch {}
			return null;
		}
		
		private AutomationEventHandler UIAeventHandler = null;
		private AutomationEventHandler UIATextSelectionChangedEventHandler = null;
		private AutomationPropertyChangedEventHandler UIAPropChangedEventHandler = null;
		
		/// <summary>
        /// Delegate for Text Changed event
        /// </summary>
		/// <param name="sender">The edit control that sent the event</param>
		/// <param name="newText">the text of the edit control</param>
		public delegate void TextChanged(UIDA_Edit sender, string newText);
		internal TextChanged TextChangedHandler = null;
		
		/// <summary>
        /// Delegate for Text Selection Changed event
        /// </summary>
		/// <param name="sender">The edit control that sent the event</param>
		/// <param name="selectedText">the selected text of the edit control</param>
		public delegate void TextSelectionChanged(UIDA_Edit sender, string selectedText);
		internal TextSelectionChanged TextSelectionChangedHandler = null;
		
		/// <summary>
        /// Attaches/detaches a handler to text changed event
        /// </summary>
		public event TextChanged TextChangedEvent
		{
			add
			{
				try
				{
					if (this.TextChangedHandler == null)
					{
						string cfid = base.uiElement.Current.FrameworkId;
						if (cfid == "Win32" || cfid == "WinForm")
						{
							UIAPropChangedEventHandler = new AutomationPropertyChangedEventHandler(
									OnUIAutomationPropChangedEvent);
									
							Automation.AddAutomationPropertyChangedEventHandler(base.uiElement, TreeScope.Element,
									UIAPropChangedEventHandler, ValuePattern.ValueProperty);
						}
						else
						{
							this.UIAeventHandler = new AutomationEventHandler(OnUIAutomationEvent);
							
							Automation.AddAutomationEventHandler(TextPattern.TextChangedEvent,
										base.uiElement, TreeScope.Element, this.UIAeventHandler);
						}
					}
					
					this.TextChangedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.TextChangedHandler -= value;
				
					if (this.TextChangedHandler == null)
					{
						string cfid = base.uiElement.Current.FrameworkId;
						if (cfid == "Win32" || cfid == "WinForm")
						{
							RemoveEventHandlerWin32();
						}
						else
						{
							RemoveEventHandler();
						}
					}
				}
				catch {}
			}
		}
		
		/// <summary>
        /// Attaches/detaches a handler to text selection changed event
        /// </summary>
		public event TextSelectionChanged TextSelectionChangedEvent
		{
			add
			{
				try
				{
					if (this.TextSelectionChangedHandler == null)
					{
						this.UIATextSelectionChangedEventHandler = new AutomationEventHandler(OnUIATextSelectionChangedEvent);
						
						Automation.AddAutomationEventHandler(TextPattern.TextSelectionChangedEvent,
									base.uiElement, TreeScope.Element, this.UIATextSelectionChangedEventHandler);
					}
					
					this.TextSelectionChangedHandler += value;
				}
				catch {}
			}
			remove
			{
				try
				{
					this.TextSelectionChangedHandler -= value;
				
					if (this.TextSelectionChangedHandler == null)
					{
						if (this.UIATextSelectionChangedEventHandler == null)
						{
							return;
						}
						
						System.Threading.Tasks.Task.Run(() => 
						{
							try
							{
								Automation.RemoveAutomationEventHandler(TextPattern.TextSelectionChangedEvent, 
									base.uiElement, this.UIATextSelectionChangedEventHandler);
								UIATextSelectionChangedEventHandler = null;
							}
							catch { }
						}).Wait(5000);
					}
				}
				catch {}
			}
		}
		
		private void RemoveEventHandlerWin32()
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
		
		private void RemoveEventHandler()
		{
			if (this.UIAeventHandler == null)
			{
				return;
			}
			
			System.Threading.Tasks.Task.Run(() => 
			{
				try
				{
					Automation.RemoveAutomationEventHandler(TextPattern.TextChangedEvent, 
						base.uiElement, this.UIAeventHandler);
					UIAeventHandler = null;
				}
				catch { }
			}).Wait(5000);
		}
		
		private void OnUIAutomationEvent(object sender, AutomationEventArgs e)
		{
			if (e.EventId == TextPattern.TextChangedEvent && this.TextChangedHandler != null)
			{
				string text = null;
				try
				{
					text = this.GetText();
				}
				catch { }
			
				this.TextChangedHandler(this, text);
			}
		}
		
		private void OnUIAutomationPropChangedEvent(object sender, AutomationPropertyChangedEventArgs e)
		{
			if (e.Property.Id == ValuePattern.ValueProperty.Id && this.TextChangedHandler != null)
			{
				string text = null;
				if (e.NewValue != null && e.NewValue is string)
				{
					text = (string)e.NewValue;
				}
				else
				{
					try
					{
						text = this.GetText();
					}
					catch { }
				}
			
				this.TextChangedHandler(this, text);
			}
		}
		
		private string previousSelectedText = "";
		private void OnUIATextSelectionChangedEvent(object sender, AutomationEventArgs e)
		{
			if (e.EventId == TextPattern.TextSelectionChangedEvent && this.TextSelectionChangedHandler != null)
			{
				string selectedText = this.GetSelectedText();
				if (selectedText != previousSelectedText)
				{
					previousSelectedText = selectedText;
					this.TextSelectionChangedHandler(this, selectedText);
				}
			}
		}
    }
}
