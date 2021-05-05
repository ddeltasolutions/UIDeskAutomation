using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Represents a RadioButton UI element.
    /// </summary>
    public class RadioButton: ElementBase
    {
        public RadioButton(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Returns the selected state.
        /// </summary>
        /// <returns>true - is selected, false otherwise</returns>
        public bool GetIsSelected()
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

        /// <summary>
        /// Selects a radio button.
        /// </summary>
        public void Select()
        {
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
            throw new Exception("RadioButton.Select() - method failed");
        }
    }
}
