using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Outlook.Enums
{
    /// <summary>
    /// Indicates the Outlook item type.
    /// </summary>
    public enum ItemType
    {
        /// <summary>
        /// An AppointmentItem object.
        /// </summary>
        AppointmentItem = 1,
        /// <summary>
        /// A ContactItem object.
        /// </summary>
        ContactItem = 2,
        /// <summary>
        /// A DistListItem object.
        /// </summary>
        DistributionListItem = 7,
        /// <summary>
        /// A JournalItem object.
        /// </summary>
        JournalItem = 4,
        /// <summary>
        /// A MailItem object.
        /// </summary>
        MailItem = 0,
        /// <summary>
        /// A MobileItem object that is a Multimedia Messaging Service (MMS) message.
        /// </summary>
        MobileItemMMS = 9,
        /// <summary>
        /// A MobileItem object that is a Short Message Service (SMS) message.
        /// </summary>
        MobileItemSMS = 8,
        /// <summary>
        /// A NoteItem object.
        /// </summary>
        NoteItem = 5,
        /// <summary>
        /// A PostItem object.
        /// </summary>
        PostItem = 6,
        /// <summary>
        /// A TaskItem object.
        /// </summary>
        TaskItem = 3
    }
}
