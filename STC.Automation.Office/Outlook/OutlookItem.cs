using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Outlook
{
    /// <summary>
    /// Parent class for MailItem to aid in CreateItem and traversal of items.  Does not correspond to any object in the Outlook object model.
    /// </summary>
    [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
    public abstract class OutlookItem : ComWrapper
    {
        private Events.ItemEvents_Sink _eventSink;

        internal OutlookItem(object itemObj)
            : base(itemObj)
        {
            _eventSink = new Events.ItemEvents_Sink(this);
        }

        /// <summary>
        /// Factory method to create and return the correct OutlookItem object (MailItem) for the given COM object.
        /// </summary>
        /// <param name="comObj"></param>
        /// <returns></returns>
        internal static OutlookItem ResolveType(object comObj)
        {
            if (ComWrapper.SupportsInterface(comObj, ComWrapper.GetMustSupport(typeof(MailItem))))
                return new MailItem(comObj);

            throw new NotImplementedException("Unknown object when attempting to resolve Outlook item type.");
        }

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_eventSink != null)
                {
                    _eventSink.Dispose();
                    _eventSink = null;
                }
            }
            
            base.Dispose(true);
        }

        /// <summary>
        /// Returns an Inspector object that represents an inspector initialized to contain the specified item. This item is internally cached and does not require manual disposal.
        /// </summary>
        /// <returns></returns>
        public Inspector Inspector
        {
            get
            {
                if (_inspector == null)
                    _inspector = new Inspector(InternalObject.GetType().InvokeMember("GetInspector", System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public, null, InternalObject, null));

                return _inspector;
            }
        }
        private Inspector _inspector;

        #region Events
        /// <summary>
        /// Occurs when an instance of the parent object is being opened in an Inspector.
        /// </summary>
        public event CanCancelEventHandler Opening
        {
            add { _eventSink._openEvent += value; }
            remove { _eventSink._openEvent -= value; }
        }

        /// <summary>
        /// Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.
        /// </summary>
        public event CanCancelEventHandler Closing
        {
            add { _eventSink._closeEvent += value; }
            remove { _eventSink._closeEvent -= value; }
        }

        /// <summary>
        /// Occurs when the user selects the Send action for an item, or when the Send method is called for the item, which is an instance of the parent object.
        /// </summary>
        public event CanCancelEventHandler Sending
        {
            add { _eventSink._sendEvent += value; }
            remove { _eventSink._sendEvent -= value; }
        }

        /// <summary>
        /// Occurs when an instance of the parent object is saved, either explicitly (for example, using the Save or SaveAs methods) or implicitly (for example, in response to a prompt when closing the item's inspector).
        /// </summary>
        public event CanCancelEventHandler Writing
        {
            add { _eventSink._writeEvent += value; }
            remove { _eventSink._writeEvent -= value; }
        }
        #endregion
    }
}
