using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using STC.Automation.Office.Common;

namespace STC.Automation.Office.Excel.Events
{
    /// <summary>
    /// Handler for the NewWorkbook event.
    /// </summary>
    /// <param name="workbook">The workbook in question. Will be disposed at the end of the event.</param>
    public delegate void NewWorkbookEventHandler(Workbook workbook);

    /// <summary>
    /// Handler for events relating to Ranges in Worksheets.
    /// </summary>
    /// <param name="sheet">The worksheet in question. Will be disposed at the end of the event.</param>
    /// <param name="target">The range in question. Will be disposed at the end of the event.</param>
    public delegate void SheetRangeEventHandler(Worksheet sheet, Range target);

    /// <summary>
    /// Handler for cancellable events relating to Ranges in Worksheets.
    /// </summary>
    /// <param name="sheet">The worksheet in question. Will be disposed at the end of the event.</param>
    /// <param name="target">The range in question. Will be disposed at the end of the event.</param>
    /// <param name="cancel">Set to true to attempt cancelling the action which caused this event to fire.</param>
    public delegate void SheetRangeCancelEventHandler(Worksheet sheet, Range target, ref bool cancel);

    /// <summary>
    /// Fields event notifications from Excel.AppEvents.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class AppEvents_Sink : ComEventSinkWrapper<IAppEvents>, IAppEvents
    {
        internal AppEvents_Sink(IConnectionPointContainer pointContainer)
            : base(pointContainer)
        {
        }

        /// <summary>
        /// Internal use only.
        /// </summary>
        public event NewWorkbookEventHandler NewWorkbookEvent;
        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void NewWorkbook(object ComWorkbook)
        {
            using (Workbook workbook = new Workbook(ComWorkbook))
            {
                if (NewWorkbookEvent != null)
                    NewWorkbookEvent(workbook);
            }
        }

        /// <summary>
        /// Internal use only.
        /// </summary>
        public event SheetRangeEventHandler SheetSelectionChangeEvent;
        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void SheetSelectionChange(object ComWorksheet, object ComRange)
        {
            using (Worksheet worksheet = new Worksheet(ComWorksheet))
            {
                using (Range range = new Range(ComRange))
                {
                    if (SheetSelectionChangeEvent != null)
                        SheetSelectionChangeEvent(worksheet, range);
                }
            }
        }

        /// <summary>
        /// Internal use only.
        /// </summary>
        public event SheetRangeCancelEventHandler SheetBeforeDoubleClickEvent;
        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void SheetBeforeDoubleClick(object ComWorksheet, object ComRange, ref bool Cancel)
        {
            using (Worksheet worksheet = new Worksheet(ComWorksheet))
            {
                using (Range range = new Range(ComRange))
                {
                    if (SheetBeforeDoubleClickEvent != null)
                        SheetBeforeDoubleClickEvent(worksheet, range, ref Cancel);
                }
            }
        }

        /// <summary>
        /// Internal use only.
        /// </summary>
        public event SheetRangeCancelEventHandler SheetBeforeRightClickEvent;
        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void SheetBeforeRightClick(object ComWorksheet, object ComRange, ref bool Cancel)
        {
            using (Worksheet worksheet = new Worksheet(ComWorksheet))
            {
                using (Range range = new Range(ComRange))
                {
                    if (SheetBeforeRightClickEvent != null)
                        SheetBeforeRightClickEvent(worksheet, range, ref Cancel);
                }
            }
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void SheetActivate(object Sh)
        {
            Marshal.ReleaseComObject(Sh);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void SheetDeactivate(object Sh)
        {
            Marshal.ReleaseComObject(Sh);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void SheetCalculate(object Sh)
        {
            Marshal.ReleaseComObject(Sh);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void SheetChange(object Sh, object RangeTarget)
        {
            Marshal.ReleaseComObject(Sh);
            Marshal.ReleaseComObject(RangeTarget);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void WorkbookOpen(object Workbook)
        {
            Marshal.ReleaseComObject(Workbook);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void WorkbookActivate(object Workbook)
        {
            Marshal.ReleaseComObject(Workbook);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void WorkbookDeactivate(object Workbook)
        {
            Marshal.ReleaseComObject(Workbook);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void WorkbookBeforeClose(object Workbook, ref bool Cancel)
        {
            Marshal.ReleaseComObject(Workbook);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void WorkbookBeforeSave(object Workbook, bool SaveAsUI, ref bool Cancel)
        {
            Marshal.ReleaseComObject(Workbook);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void WorkbookBeforePrint(object Workbook, ref bool Cancel)
        {
            Marshal.ReleaseComObject(Workbook);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void WorkbookNewSheet(object Workbook, object Sh)
        {
            Marshal.ReleaseComObject(Workbook);
            Marshal.ReleaseComObject(Sh);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void WorkbookAddinInstall(object Workbook)
        {
            Marshal.ReleaseComObject(Workbook);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void WorkbookAddinUninstall(object Workbook)
        {
            Marshal.ReleaseComObject(Workbook);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void WindowResize(object Workbook, object Window)
        {
            Marshal.ReleaseComObject(Workbook);
            Marshal.ReleaseComObject(Window);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void WindowActivate(object Workbook, object Window)
        {
            Marshal.ReleaseComObject(Workbook);
            Marshal.ReleaseComObject(Window);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void WindowDeactivate(object Workbook, object Window)
        {
            Marshal.ReleaseComObject(Workbook);
            Marshal.ReleaseComObject(Window);
        }

        /// <summary>
        /// Called by Excel.
        /// </summary>
        public void SheetFollowHyperlink(object Sh, object HyperlinkTarget)
        {
            Marshal.ReleaseComObject(Sh);
            Marshal.ReleaseComObject(HyperlinkTarget);
        }
    }
}
