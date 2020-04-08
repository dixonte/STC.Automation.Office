using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using System.Reflection;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Wraps an Excel.Worksheet object
    /// </summary>
    [WrapsCOM("Excel.Worksheet", "000208D8-0000-0000-C000-000000000046")]
    public class Worksheet : Sheet
    {
        private Range _cells;
        private Pictures _pictures;
        private Shapes _shapes;
        private Range _columns;
        private Range _rows;
        private PageSetup _pageSetup;
        private ChartObjects _chartobjects;
        private Hyperlinks _hyperlinks;
        private PivotTables _pivotTables;

        internal Worksheet(object worksheetObj)
            : base(worksheetObj)
        {
        }

        /// <summary>
        /// True if the AutoFilter drop-down arrows are currently displayed on the sheet. This property is independent of the FilterMode property. 
        /// </summary>
        public bool AutoFilterMode
        {
            get
            {
                return Convert.ToBoolean(InternalObject.GetType().InvokeMember("AutoFilterMode", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("AutoFilterMode", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns an AutoFilter object if filtering is on. 
        /// </summary>
        /// <remarks>The property returns Nothing if filtering is off.
        /// To create an AutoFilter object for a worksheet, you must turn autofiltering on for a range on the worksheet either manually or using the AutoFilter method of the Range object.</remarks>
        public AutoFilter AutoFilter
        {
            get
            {
                object obj = InternalObject.GetType().InvokeMember("AutoFilter", BindingFlags.GetProperty, null, InternalObject, null);
                if (obj == null)
                    return null;

                return new AutoFilter(obj);
            }
        }

        /// <summary>
        /// Gets a Range object starting at the origin ('A1'). This property is internally cached and does not require manual disposal.
        /// </summary>
        public Range Cells
        {
            get
            {
                if (_cells == null)
                {
                    _cells = Range("A1");
                }

                return _cells;
            }
        }

        /// <summary>
        /// The collection of ChartObject objects for this worksheet
        /// </summary>
        [Obsolete("Be cautious using this object. It was previously not documented in the Excel OLE model. It has since been changed from a Get property to a method invoke.")]
        public ChartObjects ChartObjects
        {
            get
            {
                if (_chartobjects == null)
                {
                    _chartobjects = new ChartObjects(InternalObject.GetType().InvokeMember("ChartObjects", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null));
                }

                return _chartobjects;
            }
                
        }

        /// <summary>
        /// Returns an object that represents either a single PivotTable report (a PivotTable object) or a collection of all the PivotTable reports (a PivotTables object) on a worksheet. Read-only.
        /// This object is internally cached and does not require manual disposal.
        /// </summary>
        public PivotTables PivotTables
        {
            get
            {
                if (_pivotTables == null)
                {
                    _pivotTables = new PivotTables(InternalObject.GetType().InvokeMember("PivotTables", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _pivotTables;
            }
        }

        /// <summary>
        /// Gets a Range object that represents the columns in the worksheet. This Range is internally cached and does not require manual disposal.
        /// </summary>
        public Range Columns
        {
            get
            {
                if (_columns == null)
                    _columns = new Range(InternalObject.GetType().InvokeMember("Columns", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _columns;
            }
        }

        /// <summary>
        /// Gets a Range object that represents all the rows on the specified worksheet. Read-only.
        /// </summary>
        public Range Rows
        {
            get
            {
                if (_rows == null)
                    _rows = new Range(InternalObject.GetType().InvokeMember("Rows", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _rows;
            }
        }

        /// <summary>
        /// Gets a PageSetup object that ocntains all the page setup settings for the specified object. Read-only. This pagesetup is internally cached and does not require manual disposal.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public PageSetup PageSetup
        {
            get
            {
                if (_pageSetup == null)
                    _pageSetup = new PageSetup(InternalObject.GetType().InvokeMember("PageSetup", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _pageSetup;
            }
        }

        /// <summary>
        /// Pastes the contents of the Clipboard onto the sheet, at the current selection.
        /// </summary>
        public void Paste()
        {
            InternalObject.GetType().InvokeMember("Paste", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Selects the worksheet
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public void Select()
        {
            InternalObject.GetType().InvokeMember("Select", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Protects a worksheet so that it cannot be modified
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public void Protect()
        {
            InternalObject.GetType().InvokeMember("Protect", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Removes protection from a sheet. This method has no effect if the sheet isn't protected.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public void UnProtect()
        {
            InternalObject.GetType().InvokeMember("Unprotect", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }



        /// <summary>
        /// Moves the sheet to another location in the workbook
        /// </summary>
        /// <param name="before">The sheet before which the moved sheet will be placed. You cannot specify 'before' if you specify 'after'.</param>
        /// <param name="after">The sheet after which the moved sheet will be placed. You cannot specify 'After' if you specify 'before'.</param>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public void Move(Worksheet before = null, Worksheet after = null)
        {
            if (before != null && after != null)
            {
                throw new ArgumentException("You cannot specify both 'before' and 'after' when moving a worksheet");
            }


            InternalObject.GetType().InvokeMember("Move", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { (before == null) ? System.Reflection.Missing.Value : before.InternalObject, (after == null) ? System.Reflection.Missing.Value : after.InternalObject });
        }

        /// <summary>
        /// Pastes the contents of the Clipboard onto the sheet, at the given location.
        /// </summary>
        /// <param name="destination">A Range object representing the location at which to paste.</param>
        public void Paste(Range destination)
        {
            InternalObject.GetType().InvokeMember("Paste", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { destination.InternalObject });
        }

        /// <summary>
        /// Pastes the contents of the Clipboard onto the sheet.
        /// </summary>
        /// <param name="createLink">True to establish a link to the source of the pasted data.</param>
        public void Paste(bool createLink)
        {
            InternalObject.GetType().InvokeMember("Paste", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { Missing.Value, createLink });
        }

        /// <summary>
        /// Gets the Pictures object for this worksheet. This property is internally cached and does not require manual disposal.
        /// </summary>
        [Obsolete("This property is hidden in Excel, and so should not be used.")]
        public Pictures Pictures
        {
            get
            {
                if (_pictures == null)
                {
                    _pictures = new Pictures(InternalObject.GetType().InvokeMember("Pictures", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _pictures;
            }
        }

        /// <summary>
        /// Returns a Hyperlinks collection that represents the hyperlinks for the worksheet. This property is internally cached and does not require manual disposal.
        /// </summary>
        public Hyperlinks Hyperlinks
        {
            get
            {
                if (_hyperlinks == null)
                {
                    _hyperlinks = new Hyperlinks(InternalObject.GetType().InvokeMember("Hyperlinks", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _hyperlinks;
            }
        }

        /// <summary>
        /// Gets a Range object starting at the designated cell. The returned range must be manually disposed.
        /// </summary>
        /// <param name="cell">The cell which will be the top left corner of this Range object</param>
        /// <returns>A Range object</returns>
        public Range Range(string cell)
        {
            return new Range(InternalObject.GetType().InvokeMember("Range", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { cell }));
        }

        /// <summary>
        /// Gets the Shapes object for this worksheet. This property is internally cached and does not require manual disposal.
        /// </summary>
        public Shapes Shapes
        {
            get
            {
                if (_shapes == null)
                {
                    _shapes = new Shapes(InternalObject.GetType().InvokeMember("Shapes", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _shapes;
            }
        }

        internal override void Dispose(bool disposing)
        {
            if (_cells != null)
            {
                _cells.Dispose();
                _cells = null;
            }

            if (_columns != null)
            {
                _columns.Dispose();
                _columns = null;
            }

            if (_pictures != null)
            {
                _pictures.Dispose();
                _pictures = null;
            }

            if (_shapes != null)
            {
                _shapes.Dispose();
                _shapes = null;
            }

            if (_rows != null)
            {
                _rows.Dispose();
                _rows = null;
            }

            if (_pageSetup != null)
            {
                _pageSetup.Dispose();
                _shapes = null;
            }

            if (_chartobjects != null)
            {
                _chartobjects.Dispose();
                _chartobjects = null;
            }

            if (_hyperlinks != null)
            {
                _hyperlinks.Dispose();
                _hyperlinks = null;
            }

            if (_pivotTables != null)
            {
                _pivotTables.Dispose();
                _pivotTables = null;
            }

            base.Dispose(disposing);
        }
    }
}
