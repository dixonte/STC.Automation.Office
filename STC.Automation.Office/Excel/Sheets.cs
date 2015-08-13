using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using System.IO;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Excel.Enums;
using System.Reflection;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Wraps an Excel.Sheets object
    /// </summary>
    [WrapsCOM("Excel.Sheets", "000208D7-0000-0000-C000-000000000046")]
    public class Sheets : ComWrapper
    {
        internal Sheets(object worksheetsObj)
            : base(worksheetsObj)
        {
        }

        // TODO: Fix this so it works with all sheet types (worksheet, chart, macro)
        /// <summary>
        /// Returns a single sheet from the collection. This sheet must be manually disposed.
        /// </summary>
        /// <param name="key">Sheet name</param>
        /// <returns>A Worksheet</returns>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public Sheet this[string key]
        {
            get
            {
                try
                {
                    return Sheet.ResolveType(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { key }));
                }
                catch (Exception ex)
                {
                    throw new IndexOutOfRangeException(String.Concat("Could not find worksheet '", key, "'."), ex);
                }
            }
        }

        /// <summary>
        /// Returns a single sheet from the collection. This sheet must be manually disposed.
        /// </summary>
        /// <param name="key">Sheet index</param>
        /// <returns>A Worksheet or Chart.</returns>
        public Sheet this[int key]
        {
            get
            {
                try
                {
                    return Sheet.ResolveType(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { key }));
                }
                catch (Exception ex)
                {
                    throw new IndexOutOfRangeException(String.Concat("Could not find worksheet '", key, "'."), ex);
                }
            }
        }

        /// <summary>
        /// Checks (by attempting access and catching exception) if the given worksheet name exists.
        /// </summary>
        /// <param name="name">Worksheet name</param>
        /// <returns>True if the worksheet exists</returns>
        public bool HasWorksheet(string name)
        {
            try
            {
                // TODO: enumerate worksheets instead of catching exception
                this[name].Dispose();
                return true;
            }
            catch (IndexOutOfRangeException)
            {
                return false;
            }
        }

        /// <summary>
        /// Creates a new worksheet, chart, or macro sheet. The new worksheet becomes the active sheet. The returned sheet must be manually disposed.
        /// </summary>
        /// <remarks>If Before and After are both omitted, the new sheet is inserted before the active sheet.</remarks>
        /// <param name="before">An object that specifies the sheet before which the new sheet is added.</param>
        /// <param name="after">An object that specifies the sheet after which the new sheet is added.</param>
        /// <param name="count">The number of sheets to be added. The default value is one.</param>
        /// <param name="type">Specifies the sheet type.</param>
        /// <param name="path"> If you are inserting a sheet based on an existing template, specify the path to the template (use instead of SheetType).</param>
        /// <returns>An Object value that represents the new worksheet, chart, or macro sheet.</returns>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public Sheet Add(Sheet before = null, Sheet after = null, int count = 1, SheetType type = SheetType.Worksheet, string path = null)
        {
            return Sheet.ResolveType(InternalObject.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, 
                ComArguments.Prepare(before, after, count, (String.IsNullOrEmpty(path) ? (object)(int)type : (object)path))));
        }

        /// <summary>
        /// Returns an integer value that represents the number of objects in the collection.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public int Count
        {
            get
            {
                return (int)InternalObject.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }
        }
    }
}
