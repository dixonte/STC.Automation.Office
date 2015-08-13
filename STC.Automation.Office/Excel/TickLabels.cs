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
    [WrapsCOM("Excel.TickLabels", "000208C9-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class TickLabels : ComWrapper
    {


        internal TickLabels(object interiorObj)
            : base(interiorObj)
        {
        }


        /// <summary>
        /// Returns or sets a String value that represents the format code for the object.
        /// </summary>
        public string NumberFormat
        {
            get
            {
                return InternalObject.GetType().InvokeMember("NumberFormat", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null).ToString();

            }

            set
            {
                InternalObject.GetType().InvokeMember("NumberFormat", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets a Borderobject for the series
        /// </summary>
        public Border Border
        {
            get
            {
                return new Border(InternalObject.GetType().InvokeMember("Border", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }


        /// <summary>
        /// Gets a Borderobject for the series
        /// </summary>
        public MarkerStyles MarkerStyle
        {
            get
            {
                return (MarkerStyles)InternalObject.GetType().InvokeMember("MarkerStyle", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("MarkerStyle", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }







    }
}
