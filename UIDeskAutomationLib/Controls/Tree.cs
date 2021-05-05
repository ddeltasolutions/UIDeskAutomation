using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace dDeltaSolutions.PSLib
{
    /// <summary>
    /// Class that represents a Tree control.
    /// </summary>
    public class Tree: ElementBase
    {
        public Tree(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets the tree root item.
        /// </summary>
        /// <returns>TreeItem element</returns>
        public TreeItem GetRoot()
        {
            AutomationElement rootElement = this.FindFirst(ControlType.TreeItem, null,
                false, false, true);

            if (rootElement == null)
            {
                Engine.TraceInLogFile("GetRoot() method - root element not found");

                if (Engine.ThrowExceptionsWhenSearch == true)
                {
                    throw new Exception("GetRoot() method - root element not found");
                }
                else
                {
                    return null;
                }
            }

            TreeItem root = new TreeItem(rootElement);
            return root;
        }
    }
}
