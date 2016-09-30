using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Diagnostics;
using STC.Automation.Office.Common;
using STC.Automation.Office.Excel.Events;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Wraps the Excel.Application COM object
    /// </summary>
    [WrapsCOM("Excel.Application", Application.UUID)]
    public class Application : OfficeApplication
    {
        public const string UUID = "000208D5-0000-0000-C000-000000000046";

        private Workbooks _workbooks;

        // Events
        private AppEvents_Sink _eventSink;

        /// <summary>
        /// Creates a new instance of Excel for the purposes of automation
        /// </summary>
        public Application()
            : base()
        {
            _eventSink = new AppEvents_Sink(InternalObject as IConnectionPointContainer);
        }

        internal Application(object applicationObj)
            : base(applicationObj)
        {
            _eventSink = new AppEvents_Sink(InternalObject as IConnectionPointContainer);
        }

        /// <summary>
        /// Attempts to attach to an already running Excel process.
        /// </summary>
        /// <param name="processToAttach">The Process object to which to attach.</param>
        /// <returns>An Application wrapper.</returns>
        public static Application FromProcess(Process processToAttach)
        {
            using (Window window = ComWrapper.FromProcess<Window>(processToAttach, "EXCEL7"))
            {
                if (window != null)
                {
                    return window.GetNewApplication();
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets a list of all running Excel applications from the Running Object Table... in theory. In practise, due to Excel registering in the ROT with the same name every time,
        /// what you'll is x copies of the first Excel application, where x is the number of running Excel applications. This method remains just in case future versions of Office fix this
        /// issue, and because it could be useful to know how many instances of Excel are running. Each instance should be manually disposed.
        /// </summary>
        /// <returns>A list of Excel.Application objects</returns>
        public static IList<Application> GetRunningApplications()
        {
            return Application.FromROT<Application>();
        }

        /// <summary>
        /// Returns a Range object that represents the active cell in the active window (the window on top) or in the specified window. If the window isn't displaying a worksheet, this property fails.
        /// This object must be manually disposed.
        /// </summary>
        public Range ActiveCell
        {
            get
            {
                return new Range(InternalObject.GetType().InvokeMember("ActiveCell", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Returns a Workbook object that represents the workbook in the active window (the window on top). Read-only. 
        /// Returns null if there are no windows open or if either the info window or the clipboard window is the active window.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public Workbook ActiveWorkbook
        {
            get
            {
                return new Workbook(InternalObject.GetType().InvokeMember("ActiveWorkbook", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Returns a Chart object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns Nothing.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public Chart ActiveChart
        {
            get
            {
                return new Chart(InternalObject.GetType().InvokeMember("ActiveChart", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Returns an object that represents the active sheet (the sheet on top) in the active workbook or in the specified window or workbook.
        /// Returns null if no sheet is active.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public Worksheet ActiveSheet
        {
            get
            {
                return new Worksheet(InternalObject.GetType().InvokeMember("ActiveSheet", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Provides methods for dealing with workbooks (.xls files). This object is internally cached and does not require manual disposal.
        /// </summary>
        public Workbooks Workbooks
        {
            get
            {
                if (_workbooks == null)
                {
                    _workbooks = new Workbooks(InternalObject.GetType().InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _workbooks;
            }
        }

        /// <summary>
        /// Gets the current selection as an object. Currently supported objects are Excel.Range, Excel.Shape and Excel.Picture. This object must be manually disposed.
        /// </summary>
        public ComWrapper Selection
        {
            get
            {
                object selectionCom = InternalObject.GetType().InvokeMember("Selection", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
                ComWrapper selection;

                string typeName = STC.Automation.Office.Interfaces.DispatchHelper.GetIDispatchTypeName(selectionCom as STC.Automation.Office.Interfaces.IDispatch);

                switch (typeName)
                {
                    case "Range":
                        selection = new Excel.Range(selectionCom);
                        break;

                    case "Shape":
                        selection = new Excel.Shape(selectionCom);
                        break;

                    case "Picture":
                        selection = new Excel.Picture(selectionCom);
                        break;

                    default:
                        Marshal.ReleaseComObject(selectionCom);

                        throw new NotImplementedException("Current selection TypeName is '{0}'. Automatic wrapping of this class by Selection has not yet been implemented.");
                }

                return selection;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool DisplayAlerts
        {
            get
            {
                return (bool)InternalObject.GetType().InvokeMember("DisplayAlerts", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("DisplayAlerts", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool ScreenUpdating
        {
            get
            {
                return (bool)InternalObject.GetType().InvokeMember("ScreenUpdating", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("ScreenUpdating", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns the Microsoft Excel version number.
        /// </summary>
        public Version Version
        {
            get
            {
                return new Version(InternalObject.GetType().InvokeMember("Version", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null).ToString());
            }
        }

        /// <summary>
        /// Gets or sets the visibility of the Excel program window
        /// </summary>
        public bool Visible
        {
            get
            {
                return (bool)InternalObject.GetType().InvokeMember("Visible", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Tells Excel to close itself. It may not actually close if you are still holding references to Excel objects; use of the using() clause is recommended.
        /// </summary>
        public void Quit()
        {
            InternalObject.GetType().InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        
        public object Run(string proc, params object[] args)
        {
            List<object> inArgs = new List<object>(args);
            inArgs.Insert(0, proc);

            return InternalObject.GetType().InvokeMember("Run", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, inArgs.ToArray());
        }
        
        /// <summary>
        /// Converts a measurement from centimeters to points (one point equals 0.035 centimeters)
        /// </summary>
        /// <param name="centimeters">Specifies the centimeter value to be converted to points.</param>
        /// <returns>Double - a value in points.</returns>
        [System.Obsolete("This method has not been fully tested yet and is not guaranteed to work")]
        public double CentimetersToPoints(double centimeters)
        {
            return (double)InternalObject.GetType().InvokeMember("CentimetersToPoints", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { centimeters });
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_workbooks != null)
                {
                    _workbooks.Dispose();
                    _workbooks = null;
                }

                if (_eventSink != null)
                {
                    _eventSink.Dispose();
                    _eventSink = null;
                }
            }

            base.Dispose(true);
        }

        #endregion

        #region Events
        /// <summary>
        /// Occurs when a new workbook is created.
        /// </summary>
        public event NewWorkbookEventHandler NewWorkbook
        {
            add { _eventSink.NewWorkbookEvent += value; }
            remove { _eventSink.NewWorkbookEvent -= value; }
        }

        /// <summary>
        /// Occurs when the selection changes on any worksheet (doesn't occur if the selection is on a chart sheet).
        /// </summary>
        public event SheetRangeEventHandler SheetSelectionChange
        {
            add { _eventSink.SheetSelectionChangeEvent += value; }
            remove { _eventSink.SheetSelectionChangeEvent -= value; }
        }

        /// <summary>
        /// Occurs when any worksheet is double-clicked, before the default double-click action.
        /// </summary>
        public event SheetRangeCancelEventHandler SheetBeforeDoubleClick
        {
            add { _eventSink.SheetBeforeDoubleClickEvent += value; }
            remove { _eventSink.SheetBeforeDoubleClickEvent -= value; }
        }

        /// <summary>
        /// Occurs when any worksheet is right-clicked, before the default right-click action.
        /// </summary>
        public event SheetRangeCancelEventHandler SheetBeforeRightClick
        {
            add { _eventSink.SheetBeforeRightClickEvent += value; }
            remove { _eventSink.SheetBeforeRightClickEvent -= value; }
        }
        #endregion
    }
}
