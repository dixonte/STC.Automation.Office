using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using System.IO;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Wraps an Excel.Workbooks object
    /// </summary>
    [WrapsCOM("Excel.Workbooks", "000208DB-0000-0000-C000-000000000046")]
    public class Workbooks : ComWrapper
    {
        internal Workbooks(object workbooksObj)
            : base(workbooksObj)
        {
        }

        /// <summary>
        /// Creates a new Workbook. The returned workbook must be manually disposed.
        /// </summary>
        /// <returns>A reference to the new Workbook</returns>
        public Workbook Add()
        {
            return new Workbook(InternalObject.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null));
        }

        /// <summary>
        ///  Closes all open workbooks. If there are changes in any open workbook, Microsoft Excel displays the appropriate prompts and dialog boxes for saving changes.
        /// </summary>
        /// <returns>A reference to the new Workbook</returns>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public Workbook Close()
        {
            return new Workbook(InternalObject.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null));
        }

        /// <summary>
        /// Opens an existing workbook. The returned workbook must be manually disposed.
        /// </summary>
        /// <param name="filename">The full path to the workbook file (.xls)</param>
        /// <returns>A reference to the Workbook</returns>
        /// <exception cref="FileNotFoundException">If the file does not exist</exception>
        public Workbook Open(string filename)
        {
            return Open(filename, null);
        }

        /// <summary>
        /// Opens an existing workbook. The returned workbook must be manually disposed.
        /// </summary>
        /// <param name="filename">The full path to the workbook file (.xls)</param>
        /// <param name="readOnly">Open the workbook as read only</param>
        /// <returns>A reference to the Workbook</returns>
        /// <exception cref="FileNotFoundException">If the file does not exist</exception>
        public Workbook Open(string filename, bool? readOnly)
        {
            if (!File.Exists(filename))
                throw new FileNotFoundException();

            return new Workbook(InternalObject.GetType().InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, 
                new object[] { 
                    filename, 
                    System.Reflection.Missing.Value, 
                    (readOnly.HasValue ? (object)readOnly : System.Reflection.Missing.Value) 
                }));
        }

        /// <summary>
        /// Index the workbooks
        /// </summary>
        /// <param name="id">workbook index</param>
        /// <returns></returns>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public Workbook this[int id]
        {
            get
            {
                return new Workbook(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { id }));
            }
        }

        /// <summary>
        /// Gets a generic IEnumerator of Workbook objects.
        /// </summary>
        /// <returns>IEnumerator&lt;Workbook&gt;</returns>
        public IEnumerator<Workbook> GetEnumerator()
        {
            return new ComIEnumeratorWrapper<Workbook>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        }
    }
}
