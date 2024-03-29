﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a Hyper Link control.
    /// </summary>
    public class UIDA_HyperLink: ElementBase
    {
        public UIDA_HyperLink(AutomationElement el)
        {
            this.uiElement = el;
        }
        
        /// <summary>
        /// Accesses the link
        /// </summary>
        public void AccessLink()
        {
            base.Invoke();
        }
    }
}
