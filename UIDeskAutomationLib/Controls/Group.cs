using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// This class represents a Group.
    /// </summary>
    public class UIDA_Group : ElementBase
    {
        /// <summary>
        /// Creates a UIDA_Group using an AutomationElement
        /// </summary>
        /// <param name="el">UI Automation AutomationElement</param>
        public UIDA_Group(AutomationElement el)
        {
			this.uiElement = el;
		}

        /// <summary>
        /// Expands a group.
        /// </summary>
        public void Expand()
        {
            ExpandCollapsePattern expandCollapsePattern = this.GetExpandCollapsePattern();

            if (expandCollapsePattern == null)
            {
                Engine.TraceInLogFile(
                    "UIDA_Group.Expand() - ExpandCollapsePattern not supported");

                throw new Exception(
                    "UIDA_Group.Expand() - ExpandCollapsePattern not supported");
            }

            try
            {
                if (expandCollapsePattern.Current.ExpandCollapseState !=
                    ExpandCollapseState.Expanded)
                {
                    expandCollapsePattern.Expand();
                }
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("UIDA_Group.Expand() error: " + ex.Message);
                throw new Exception("UIDA_Group.Expand() error: " + ex.Message);
            }
        }

        /// <summary>
        /// Collapses a group.
        /// </summary>
        public void Collapse()
        {
            ExpandCollapsePattern expandCollapsePattern = this.GetExpandCollapsePattern();
            if (expandCollapsePattern == null)
            {
                Engine.TraceInLogFile(
                    "UIDA_Group.Collapse() - ExpandCollapsePattern not supported");

                throw new Exception(
                    "UIDA_Group.Collapse() - ExpandCollapsePattern not supported");
            }

            try
            {
                if (expandCollapsePattern.Current.ExpandCollapseState != ExpandCollapseState.Collapsed)
                {
                    expandCollapsePattern.Collapse();
                }
            }
            catch (Exception ex)
            {
                Engine.TraceInLogFile("UIDA_Group.Collapse() error: " + ex.Message);
                throw new Exception("UIDA_Group.Collapse() error: " + ex.Message);
            }
        }

        private ExpandCollapsePattern GetExpandCollapsePattern()
        {
            object expandCollapsePatternObj = null;
            if (this.uiElement.TryGetCurrentPattern(ExpandCollapsePattern.Pattern,
                out expandCollapsePatternObj) == true)
            {
                ExpandCollapsePattern expandCollapsePattern = expandCollapsePatternObj as ExpandCollapsePattern;
                return expandCollapsePattern;
            }

            return null;
        }
    }
}
