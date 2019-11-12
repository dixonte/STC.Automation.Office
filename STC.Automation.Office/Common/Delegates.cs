using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Common
{
    /// <summary>
    /// Common event handler delegate for events that allow an action to be cancelled.
    /// </summary>
    /// <param name="sender">The object receiving the event.</param>
    /// <param name="cancel">Set to true to cancel the action.</param>
    public delegate void CanCancelEventHandler(object sender, ref bool cancel);
}
