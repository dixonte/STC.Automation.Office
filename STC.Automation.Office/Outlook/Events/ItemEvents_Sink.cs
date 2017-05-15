using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace STC.Automation.Office.Outlook.Events
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class ItemEvents_Sink : ComEventSinkWrapper<IItemEvents>, IItemEvents
    {
        internal ItemEvents_Sink(ComWrapper parent)
            : base(parent)
        {
        }

        internal CanCancelEventHandler _openEvent;
        public void Open(ref bool Cancel)
        {
            _openEvent?.Invoke(Parent.Target, ref Cancel);
        }

        public void CustomAction([MarshalAs(UnmanagedType.IUnknown)] object Action, [MarshalAs(UnmanagedType.IUnknown)] object Response, ref bool Cancel)
        {
            Marshal.ReleaseComObject(Action);
            Marshal.ReleaseComObject(Response);
        }

        public void CustomPropertyChange([MarshalAs(UnmanagedType.BStr)] string Name)
        {
        }

        public void Forward([MarshalAs(UnmanagedType.IUnknown)] object Forward, ref bool Cancel)
        {
            Marshal.ReleaseComObject(Forward);
        }

        internal CanCancelEventHandler _closeEvent;
        public void Close(ref bool Cancel)
        {
            _closeEvent?.Invoke(Parent.Target, ref Cancel);
        }

        public void PropertyChange([MarshalAs(UnmanagedType.BStr)] string Name)
        {
        }

        public void Read()
        {
        }

        public void Reply([MarshalAs(UnmanagedType.IUnknown)] object Response, ref bool Cancel)
        {
            Marshal.ReleaseComObject(Response);
        }

        public void ReplyAll([MarshalAs(UnmanagedType.IUnknown)] object Response, ref bool Cancel)
        {
            Marshal.ReleaseComObject(Response);
        }

        internal CanCancelEventHandler _sendEvent;
        public void Send(ref bool Cancel)
        {
            _sendEvent?.Invoke(Parent.Target, ref Cancel);
        }

        internal CanCancelEventHandler _writeEvent;
        public void Write(ref bool Cancel)
        {
            _writeEvent?.Invoke(Parent.Target, ref Cancel);
        }

        public void BeforeCheckNames(ref bool Cancel)
        {
        }

        public void AttachmentAdd([MarshalAs(UnmanagedType.IUnknown)] object Attachment)
        {
            Marshal.ReleaseComObject(Attachment);
        }

        public void AttachmentRead([MarshalAs(UnmanagedType.IUnknown)] object Attachment)
        {
            Marshal.ReleaseComObject(Attachment);
        }

        public void BeforeAttachmentSave([MarshalAs(UnmanagedType.IUnknown)] object Attachment, ref bool Cancel)
        {
            Marshal.ReleaseComObject(Attachment);
        }
    }
}
