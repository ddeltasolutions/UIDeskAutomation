﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
    /// <summary>
    /// Represents a StatusBar control.
    /// </summary>
    public class UIDA_StatusBar: ElementBase
    {
        public UIDA_StatusBar(AutomationElement el)
        {
            this.uiElement = el;
        }
    }
}
