using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;

namespace STC.Automation.Office.Excel.Events
{
    /// <summary>
    /// Imports events for Excel.Application
    /// </summary>
    [ComImport, Guid("00024413-0000-0000-C000-000000000046")]
    [TypeLibType(TypeLibTypeFlags.FCanCreate | TypeLibTypeFlags.FPreDeclId), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IAppEvents
    {
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x61d)]
        void NewWorkbook([In, MarshalAs(UnmanagedType.Interface)] object Workbook);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x616)]
        void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] object RangeTarget);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x617)]
        void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] object RangeTarget, [In] ref bool Cancel);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x618)]
        void SheetBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] object RangeTarget, [In] ref bool Cancel);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x619)]
        void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object Sh);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x61a)]
        void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object Sh);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x61b)]
        void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object Sh);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x61c)]
        void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] object RangeTarget);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x61f)]
        void WorkbookOpen([In, MarshalAs(UnmanagedType.Interface)] object Workbook);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x620)]
        void WorkbookActivate([In, MarshalAs(UnmanagedType.Interface)] object Workbook);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x621)]
        void WorkbookDeactivate([In, MarshalAs(UnmanagedType.Interface)] object Workbook);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x622)]
        void WorkbookBeforeClose([In, MarshalAs(UnmanagedType.Interface)] object Workbook, [In] ref bool Cancel);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x623)]
        void WorkbookBeforeSave([In, MarshalAs(UnmanagedType.Interface)] object Workbook, [In] bool SaveAsUI, [In] ref bool Cancel);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x624)]
        void WorkbookBeforePrint([In, MarshalAs(UnmanagedType.Interface)] object Workbook, [In] ref bool Cancel);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x625)]
        void WorkbookNewSheet([In, MarshalAs(UnmanagedType.Interface)] object Workbook, [In, MarshalAs(UnmanagedType.IDispatch)] object Sh);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x626)]
        void WorkbookAddinInstall([In, MarshalAs(UnmanagedType.Interface)] object Workbook);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x627)]
        void WorkbookAddinUninstall([In, MarshalAs(UnmanagedType.Interface)] object Workbook);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x612)]
        void WindowResize([In, MarshalAs(UnmanagedType.Interface)] object Workbook, [In, MarshalAs(UnmanagedType.Interface)] object Window);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x614)]
        void WindowActivate([In, MarshalAs(UnmanagedType.Interface)] object Workbook, [In, MarshalAs(UnmanagedType.Interface)] object Window);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x615)]
        void WindowDeactivate([In, MarshalAs(UnmanagedType.Interface)] object Workbook, [In, MarshalAs(UnmanagedType.Interface)] object Window);
        /// <summary>
        /// Called by Excel.
        /// </summary>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x73e)]
        void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] object HyperlinkTarget);
    }
}
