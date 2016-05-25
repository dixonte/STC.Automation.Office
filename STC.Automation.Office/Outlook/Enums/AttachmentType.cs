using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Outlook.Enums
{
    public enum AttachmentType
    {
        /// <summary>
        /// This value is no longer supported since Microsoft Outlook 2007. Use olByValue to attach a copy of a file in the file system.
        /// </summary>
        ByReference = 4,
        /// <summary>
        /// The attachment is a copy of the original file and can be accessed even if the original file is removed.
        /// </summary>
        ByValue = 1,
        /// <summary>
        /// The attachment is an Outlook message format file (.msg) and is a copy of the original message.
        /// </summary>
        Embeddeditem = 5,
        /// <summary>
        /// The attachment is an OLE document.
        /// </summary>
        OLE = 6
    }
}
