using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using STC.Automation.Office.Outlook.Enums;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Outlook
{
    /// <summary>
    /// Wraps an Outlook.MailItem object
    /// </summary>
    [WrapsCOM("Outlook.MailItem", "00063034-0000-0000-C000-000000000046")]
    public class MailItem : OutlookItem
    {
        private Recipients _recipients;
        private Attachments _attachments;

        internal MailItem(object namespaceObj)
            : base(namespaceObj)
        {
        }

        /// <summary>
        /// Returns an Attachments object that represents all the attachments for the specified item. This object is internally cached and does not require manual disposal..
        /// </summary>
        public Attachments Attachments
        {
            get
            {
                if (_attachments == null)
                {
                    _attachments = new Attachments(InternalObject.GetType().InvokeMember("Attachments", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _attachments;
            }
        }

        /// <summary>
        /// Provides methods for dealing with recipients. This object is internally cached and does not require manual disposal.
        /// </summary>
        public Recipients Recipients
        {
            get
            {
                if (_recipients == null)
                {
                    _recipients = new Recipients(InternalObject.GetType().InvokeMember("Recipients", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _recipients;
            }
        }

        /// <summary>
        /// Returns or sets a String indicating the subject for the Outlook item. Read/write.
        /// </summary>
        public string Subject
        {
            get
            {
                return (string)InternalObject.GetType().InvokeMember("Subject", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }
            set
            {
                InternalObject.GetType().InvokeMember("Subject", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns or sets a String representing the clear-text body of the Outlook item. Read/write.
        /// </summary>
        public string Body
        {
            get
            {
                return (string)InternalObject.GetType().InvokeMember("Body", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }
            set
            {
                InternalObject.GetType().InvokeMember("Body", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns or sets a String indicating the subject for the Outlook item. Read/write.
        /// </summary>
        public Importance Importance
        {
            get
            {
                return (Importance)InternalObject.GetType().InvokeMember("Importance", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }
            set
            {
                InternalObject.GetType().InvokeMember("Importance", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Displays a new Explorer object for the MailItem.
        /// </summary>
        public void Display()
        {
            InternalObject.GetType().InvokeMember("Display", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { });
        }

        /// <summary>
        /// Sends the e-mail message.
        /// </summary>
        public void Send()
        {
            InternalObject.GetType().InvokeMember("Send", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { });
        }

        /// <summary>
        /// Saves the Microsoft Outlook item to the current folder or, if this is a new item, to the Outlook default folder for the item type.
        /// </summary>
        public void Save()
        {
            InternalObject.GetType().InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { });
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_attachments != null)
                {
                    _attachments.Dispose();
                    _attachments = null;
                }
                if (_recipients != null)
                {
                    _recipients.Dispose();
                    _recipients = null;
                }
            }

            base.Dispose(true);
        }

        #endregion
    }
}
