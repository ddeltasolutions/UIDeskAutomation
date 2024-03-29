﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Class that represents a Tree control.
    /// </summary>
    public class UIDA_Tree: ElementBase
    {
        public UIDA_Tree(AutomationElement el)
        {
            this.uiElement = el;
        }

        /// <summary>
        /// Gets the tree root item.
        /// </summary>
        /// <returns>UIDA_TreeItem element</returns>
        public UIDA_TreeItem GetRoot()
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

            UIDA_TreeItem root = new UIDA_TreeItem(rootElement);
            return root;
        }
    }
}
