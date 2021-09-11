using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a RadioButton UI element.
    /// </summary>
    public class UIDA_RadioButton: ElementBase
    {
        public UIDA_RadioButton(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets the selected state of the radio button.
        /// </summary>
        /// <returns>true - is selected, false otherwise</returns>
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

					if (selectionItemPattern == null)
					{
						Engine.TraceInLogFile("RadioButton.GetIsSelected() - method failed");
						throw new Exception("RadioButton.GetIsSelected() - method failed");
					}

					bool isSelected = false;

					try
					{
						isSelected = selectionItemPattern.Current.IsSelected;

						return isSelected;
					}
					catch { }
				}

				Engine.TraceInLogFile("RadioButton.GetIsSelected() - method failed");
				throw new Exception("RadioButton.GetIsSelected() - method failed");
			}
        }

        /// <summary>
        /// Selects a radio button.
        /// </summary>
        public void Select()
        {
			this.Click();
			//Engine.GetInstance().Sleep(100);
			
			/*if (this.uiElement.Current.FrameworkId == "WPF")
			{
				// for WPF use mouse click because SelectionItemPattern.Select() is not calling the radio button's Click event handler
				this.Click();
				return;
			}
			
            object selectionItemPatternObj = null;

            if (this.uiElement.TryGetCurrentPattern(SelectionItemPattern.Pattern,
                out selectionItemPatternObj) == true)
            {
                SelectionItemPattern selectionItemPattern =
                    selectionItemPatternObj as SelectionItemPattern;

                if (selectionItemPattern == null)
                {
                    Engine.TraceInLogFile("RadioButton.Select() - method failed");
                    throw new Exception("RadioButton.Select() - method failed");
                }

                try
                {
                    selectionItemPattern.Select();
                    return;
                }
                catch { }
            }

            Engine.TraceInLogFile("RadioButton.Select() - method failed");
            throw new Exception("RadioButton.Select() - method failed"); */
        }
		
		/// <summary>
        /// Gets the text of the radio button
        /// </summary>
		public string Text
		{
			get
			{
				return this.GetText();
			}
		}
    }
}
