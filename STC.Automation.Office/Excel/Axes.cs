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
    /// Represents the interior of an object.
    /// </summary>
    [WrapsCOM("Excel.Axes", "0002085B-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class Axes : ComWrapper
    {

        internal Axes(object interiorObj)
            : base(interiorObj)
        {
        }


        /// <summary>
        /// Index the Axes collection to get an axis
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public Axis this[int key]
        {
            get
            {
                try
                {
                    return new Axis(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { key }));
                }
                catch (Exception ex)
                {
                    throw new IndexOutOfRangeException(String.Concat("Could not find axis '", key, "'."), ex);
                }
            }
        }

        /// <summary>
        /// Index the Axes collection to get an axis
        /// </summary>
        /// <param name="type">Specifies the axis to return.</param>
        /// <returns></returns>
        public Axis this[AxisType type]
        {
            get
            {
                try
                {
                    return new Axis(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { (int)type }));
                }
                catch (Exception ex)
                {
                    throw new IndexOutOfRangeException(String.Concat("Could not find axis '", type, "'."), ex);
                }
            }
        }

        /// <summary>
        /// Index the Axes collection to get an axis
        /// </summary>
        /// <param name="type">Specifies the axis to return.</param>
        /// <param name="group">Specifies the axis group. If this argument is omitted, the primary group is used. 3-D charts have only one axis group.</param>
        /// <returns></returns>
        public Axis this[AxisType type, AxisGroup group]
        {
            get
            {
                try
                {
                    return new Axis(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { (int)type, (int)group }));
                }
                catch (Exception ex)
                {
                    throw new IndexOutOfRangeException(String.Concat("Could not find axis '", type, "', '", group, "'."), ex);
                }
            }
        }

    }
}
