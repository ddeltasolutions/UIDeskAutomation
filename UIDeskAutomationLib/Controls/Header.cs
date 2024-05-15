using System;
using System.Collections.Generic;
using System.Windows.Automation;

namespace UIDeskAutomationLib
{
	/// <summary>
	/// This class represents a header.
	/// </summary>
	public class UIDA_Header: ElementBase
	{
		/// <summary>
		/// Creates a UIDA_Header using an AutomationElement
		/// </summary>
		/// <param name="el">UI Automation AutomationElement</param>
		public UIDA_Header(AutomationElement el)
		{
			this.uiElement = el;
		}

		/// <summary>
		/// Gets the header items in a header.
		/// </summary>
		public UIDA_HeaderItem[] Items
		{
			get
			{
				List<AutomationElement> items = this.FindAll(ControlType.HeaderItem,
					null, false, false, true);

				List<UIDA_HeaderItem> headerItems = new List<UIDA_HeaderItem>();

				foreach (AutomationElement item in items)
				{ 
					UIDA_HeaderItem headerItem = new UIDA_HeaderItem(item);
					headerItems.Add(headerItem);
				}

				return headerItems.ToArray();
			}
		}
	}
}