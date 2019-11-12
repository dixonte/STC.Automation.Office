using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace STC.Automation.Office.Outlook.Events
{
    /// <summary>
    /// Imports events for Excel.Application
    /// </summary>
    [ComImport, Guid("0006303A-0000-0000-C000-000000000046")]
    [TypeLibType(TypeLibTypeFlags.FCanCreate | TypeLibTypeFlags.FPreDeclId), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IItemEvents
    {
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f003)]
        void Open(ref bool Cancel);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f006)]
        void CustomAction(/*[in] IDispatch* */ [MarshalAs(UnmanagedType.IUnknown)] object Action, /*[in] IDispatch* */ [MarshalAs(UnmanagedType.IUnknown)] object Response, ref bool Cancel);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f008)]
        void CustomPropertyChange([MarshalAs(UnmanagedType.BStr)] string Name);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f468)]
        void Forward(/*[in] IDispatch* */ [MarshalAs(UnmanagedType.IUnknown)] object Forward, ref bool Cancel);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f004)]
        void Close(ref bool Cancel);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f009)]
        void PropertyChange([MarshalAs(UnmanagedType.BStr)] string Name);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f001)]
        void Read();

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f466)]
        void Reply(/*[in] IDispatch* */ [MarshalAs(UnmanagedType.IUnknown)] object Response, ref bool Cancel);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f467)]
        void ReplyAll(/*[in] IDispatch* */ [MarshalAs(UnmanagedType.IUnknown)] object Response, ref bool Cancel);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f005)]
        void Send(ref bool Cancel);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f002)]
        void Write(ref bool Cancel);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f00a)]
        void BeforeCheckNames(ref bool Cancel);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f00b)]
        void AttachmentAdd(/*[in] Attachment* */ [MarshalAs(UnmanagedType.IUnknown)] object Attachment);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f00c)]
        void AttachmentRead(/*[in] Attachment* */ [MarshalAs(UnmanagedType.IUnknown)] object Attachment);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x0000f00d)]
        void BeforeAttachmentSave(/*[in] Attachment* */ [MarshalAs(UnmanagedType.IUnknown)] object Attachment, ref bool Cancel);
    }
}
