using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Outlook.Enums
{
    /// <summary>
    /// Indicates the recipient type for the Item.
    /// </summary>
    public enum MailRecipientType
    {
        /// <summary>
        /// The recipient is specified in the BCC property of the Item.
        /// </summary>
        BCC = 3,
        /// <summary>
        /// The recipient is specified in the CC property of the Item.
        /// </summary>
        CC = 2,
        /// <summary>
        /// Originator (sender) of the Item.
        /// </summary>
        Originator = 0,
        /// <summary>
        /// The recipient is specified in the To property of the Item.
        /// </summary>
        To = 1
    }
}
