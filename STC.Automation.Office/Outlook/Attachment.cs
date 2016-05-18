using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Outlook
{
    /// <summary>
    /// Represents a document or link to a document contained in an Outlook item.
    /// </summary>
    [WrapsCOM("Outlook.Attachment", "00063007-0000-0000-C000-000000000046")]
    public class Attachment : ComWrapper
        //where T : struct, IConvertible // Enum
    {
        internal Attachment(object attachmentObj)
            : base(attachmentObj)
        {
        }
    }
}
