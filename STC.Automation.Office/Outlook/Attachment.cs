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
        /// <summary>
        /// Usage: this.PropertyAccessor.SetProperty(PR_ATTACH_MIME_TAG, "image/jpeg")
        /// </summary>
        public const string PR_ATTACH_MIME_TAG = "http://schemas.microsoft.com/mapi/proptag/0x370E001E";

        /// <summary>
        /// Usage:
        /// this.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, "image001.jpg")
        /// mailItem.HTMLBody = "<img src=cid:image001.jpg />"
        /// </summary>
        public const string PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001E";

        /// <summary>
        /// Usage: this.PropertyAccessor.SetProperty(PR_ATTACHMENT_HIDDEN, true)
        /// </summary>
        public const string PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B";

        private PropertyAccessor _propertyAccessor;

        internal Attachment(object attachmentObj)
            : base(attachmentObj)
        {
        }

        /// <summary>
        /// Returns a PropertyAccessor object that supports creating, getting, setting, and deleting properties of the parent Attachment object. This object is internally cached and does not require manual disposal.
        /// </summary>
        public PropertyAccessor PropertyAccessor
        {
            get
            {
                if (_propertyAccessor == null)
                {
                    _propertyAccessor = new PropertyAccessor(InternalObject.GetType().InvokeMember("PropertyAccessor", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _propertyAccessor;
            }
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_propertyAccessor != null)
                {
                    _propertyAccessor.Dispose();
                    _propertyAccessor = null;
                }
            }

            base.Dispose(true);
        }

        #endregion
    }
}
