using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Outlook.Enums;

namespace STC.Automation.Office.Outlook
{
    /// <summary>
    /// Contains a set of Attachment objects that represent the attachments in an Outlook item.
    /// </summary>
    [WrapsCOM("Outlook.Attachments", "0006303C-0000-0000-C000-000000000046")]
    public class Attachments : ComWrapper
    {
        internal Attachments(object attachmentsObj)
            : base(attachmentsObj)
        {
        }

        /// <summary>
        /// Creates a new attachment in the Attachments collection. The returned Attachment must be manually disposed.
        /// </summary>
        /// <remarks>When an Attachment is added to the Attachments collection of an item, the Type property of the Attachment will always return olOLE (6) until the item is saved. To ensure consistent results, always save an item before adding or removing objects in the Attachments collection.</remarks>
        /// <param name="filepath">The source of the attachment. This can be a file (represented by the full file system path with a file name) or an Outlook item that constitutes the attachment.</param>
        /// <param name="type">The type of the attachment.</param>
        /// <param name="position">This parameter applies only to e-mail messages using the Rich Text format: it is the position where the attachment should be placed within the body text of the message. A value of 1 for the Position parameter specifies that the attachment should be positioned at the beginning of the message body. A value 'n' greater than the number of characters in the body of the e-mail item specifies that the attachment should be placed at the end. A value of 0 makes the attachment hidden.</param>
        /// <param name="displayName">This parameter applies only if the mail item is in Rich Text format and Type is set to olByValue: the name is displayed in an Inspector object for the attachment or when viewing the properties of the attachment. If the mail item is in Plain Text or HTML format, then the attachment is displayed using the file name in the Source parameter.</param>
        /// <returns>An Attachment object that represents the new attachment.</returns>
        public Attachment Add(string filepath) //, AttachmentType type = AttachmentType.OLE, long position = 0, string displayName = null)
        {
            var missing = System.Reflection.Missing.Value;

            return new Attachment(InternalObject.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { filepath, missing, missing, missing }));
        }

        /// <summary>
        /// Index the Recipients collection to get a recipient
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public Attachment this[int key]
        {
            get
            {
                try
                {
                    return new Attachment(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { key }));
                }
                catch (Exception ex)
                {
                    throw new IndexOutOfRangeException(String.Concat("Could not find series '", key, "'."), ex);
                }
            }
        }


        /// <summary>
        /// Returns an integer value that represents the number of objects in the collection.
        /// </summary>
        public int Count
        {
            get
            {
                return (int)InternalObject.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }
        }
    }
}
