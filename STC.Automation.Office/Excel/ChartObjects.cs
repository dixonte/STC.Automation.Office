using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Excel.Enums;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// The ChartObject object acts as a container for a Chart object. Properties and methods for the ChartObject object control the appearance and size of the embedded chart on the worksheet. 
    /// The ChartObject object is a member of the ChartObjects collection. The ChartObjects collection contains all the embedded charts on a single sheet.
    /// </summary>
    [WrapsCOM("Excel.ChartObjects", "000208D0-0000-0000-C000-000000000046")]
    public class ChartObjects : ComWrapper
    {

        internal ChartObjects(object interiorObj)
            : base(interiorObj)
        {
        }

        /// <summary>
        /// Creates a new embedded chart.
        /// </summary>
        public ChartObject Add(decimal left, decimal top, decimal width, decimal height)
        {
            return new ChartObject(InternalObject.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { left, top, width, height }));
        }

        /// <summary>
        /// Returns a single object from a collection.
        /// </summary>
        /// <param name="index">The index number for the object.</param>
        /// <returns></returns>
        public ChartObject Item(int index)
        {
            return new ChartObject(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { index }));
        }

        /// <summary>
        /// Returns a single object from a collection.
        /// </summary>
        /// <param name="name">The index name for the object.</param>
        /// <returns></returns>
        public ChartObject Item(string name)
        {
            return new ChartObject(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { name }));
        }

    }
}
