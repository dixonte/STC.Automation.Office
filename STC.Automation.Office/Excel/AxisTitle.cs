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
    [WrapsCOM("Excel.AxisTitle", "0002084A-0000-0000-C000-000000000046")]
    public class AxisTitle : ComWrapper
    {


        internal AxisTitle(object interiorObj)
            : base(interiorObj)
        {
        }


        /// <summary>
        /// Returns or sets a String value that represents the format code for the object.
        /// </summary>
        public string Text
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Text", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null).ToString();
            }

            set
            {
                InternalObject.GetType().InvokeMember("Text", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
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
