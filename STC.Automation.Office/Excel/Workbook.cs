using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Wraps an Excel.Workbook object
    /// </summary>
    [WrapsCOM("Excel.Workbook", "000208DA-0000-0000-C000-000000000046")]
    public class Workbook : ComWrapper
    {
        private Sheets _worksheets;
        private Sheets _sheets;
        private PivotCaches _pivotCaches;

        internal Workbook(object workbookObj)
            : base(workbookObj)
        {
        }

        /// <summary>
        /// Retrieves all open Excel workbooks from the Running Object Table. Each instance must be manually disposed.
        /// </summary>
        /// <returns>A list of workbooks.</returns>
        public static IList<Workbook> GetAllOpen()
        {
            return Application.FromROT<Workbook>();
        }

        /// <summary>
        /// Gets a reference to the currently active worksheet (tab). The returned worksheet must be manually disposed.
        /// </summary>
        public Worksheet ActiveSheet
        {
            get
            {
                var obj = InternalObject.GetType().InvokeMember("ActiveSheet", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
                if (obj == null)
                    return null;

                return new Worksheet(obj);
            }
        }

        /// <summary>
        /// Returns a Chart object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns null.  The returned Chart must be manually disposed.
        /// </summary>
        public Chart ActiveChart
        {
            get
            {
                var obj = InternalObject.GetType().InvokeMember("ActiveChart", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
                if (obj == null)
                    return null;

                return new Chart(obj);
            }
        }

        /// <summary>
        /// Returns a Sheets collection that represents all the sheets in the specified workbook. Read-only Sheets object. This object is internally cached and does not require manual disposal.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public Sheets Sheets
        {
            get
            {
                if (_sheets == null)
                {
                     _sheets = new Sheets(InternalObject.GetType().InvokeMember("Sheets", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _sheets;
            }
        }

        /// <summary>
        /// Returns a PivotCaches collection that represents all the PivotTable caches in the specified workbook. Read-only.
        /// NOTE: Microsoft treats this as a method, rather than a property. We are treating it like a property and caching its result
        /// </summary>
        public PivotCaches PivotCaches
        {
            get
            {
                if (_pivotCaches == null)
                {
                    _pivotCaches = new PivotCaches(InternalObject.GetType().InvokeMember("PivotCaches", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null));
                }

                return _pivotCaches;
            }
        }


        /// <summary>
        /// Provides methods for dealing with worksheets in this workbook. This object is internally cached and does not require manual disposal.
        /// </summary>
        public Sheets Worksheets
        {
            get
            {
                if (_worksheets == null)
                {
                    _worksheets = new Sheets(InternalObject.GetType().InvokeMember("Worksheets", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _worksheets;
            }
        }

        /// <summary>
        /// Returns true if the object has been opened as read-only.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public bool ReadOnly
        {
            get
            {
                return (bool)InternalObject.GetType().InvokeMember("ReadOnly", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }
        }


        /// <summary>
        /// Saves the workbook.
        /// </summary>
        public void Save()
        {
            InternalObject.GetType().InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Saves the workbook under a new filename.
        /// </summary>
        /// <param name="filename">The filename under which to save the workbook</param>
        public void SaveAs(string filename)
        {
            SaveAs(filename, Enums.FileFormat.WorkbookNormal, Enums.SaveAsAccessMode.NoChange);
        }

        /// <summary>
        /// Saves the workbook under a new filename.
        /// </summary>
        /// <param name="filename">The filename under which to save the workbook</param>
        /// <param name="fileFormat">The format in which to save the workbook</param>
        /// <param name="saveAsAccessMode">The file access mode to use when saving</param>
        public void SaveAs(string filename, Enums.FileFormat fileFormat, Enums.SaveAsAccessMode saveAsAccessMode)
        {
            var missing = System.Reflection.Missing.Value;

            InternalObject.GetType().InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { filename, fileFormat, missing, missing,
                false, false, saveAsAccessMode });
        }

        /// <summary>
        /// Saves the workbook under a new filename.
        /// </summary>
        /// <param name="Type">Can be either xlTypePDF or xlTypeXPS.</param>
        /// <param name="Filename">A string that indicates the name of the file to be saved. You can include a full path or Excel saves the file in the current folder.</param>
        public void ExportAsFixedFormat(Enums.FixedFormatType Type, string Filename )
        {
            var missing = System.Reflection.Missing.Value;

            InternalObject.GetType().InvokeMember("ExportAsFixedFormat", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { Type, Filename });
        }

        /// <summary>
        /// Closes the workbook.
        /// </summary>
        public void Close()
        {
            Close(null);
        }

        /// <summary>
        /// Closes the workbook.
        /// </summary>
        /// <param name="saveChanges">Saves changes if true, abandons them if false, and asks the user if null</param>
        public void Close(bool? saveChanges)
        {
            List<object> parms = new List<object>();
            if (saveChanges != null)
                parms.Add(saveChanges.Value);

            InternalObject.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, parms.ToArray());

            Dispose(true);
        }

        /// <summary>
        /// Returns the name of the object, including its path on disk, as a string. Read-only String.
        /// </summary>
        public string FullName
        {
            get
            {
                return InternalObject.GetType().InvokeMember("FullName", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null) as string;
            }
        }


        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_worksheets != null)
                {
                    _worksheets.Dispose();
                    _worksheets = null;
                }

                if (_sheets != null)
                {
                    _sheets.Dispose();
                    _sheets = null;
                }

                if (_pivotCaches != null)
                {
                    _pivotCaches.Dispose();
                    _pivotCaches = null;
                }
            }

            base.Dispose(true);
        }

        #endregion
    }
}
