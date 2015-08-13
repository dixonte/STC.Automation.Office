using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Wraps an Excel Shape object. Could contain a picture, chart, label, callout, etc.
    /// </summary>
    [WrapsCOM("Excel.Shape", "00024439-0000-0000-C000-000000000046")]
    public class Shape : ComWrapper
    {
        internal Shape(object shapeObj)
            : base(shapeObj)
        {
        }

        /// <summary>
        /// Gets or sets a String value representing the name of the Shape.
        /// </summary>
        public string Name
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null).ToString();
            }

            set
            {
                InternalObject.GetType().InvokeMember("Name", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns or sets a value that represents the width, in points, of the object.
        /// </summary>
        public double Width
        {
            get
            {
                return Convert.ToDouble(InternalObject.GetType().InvokeMember("Width", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Width", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns or sets a value that represents the height, in points, of the object.
        /// </summary>
        public double Height
        {
            get
            {
                return Convert.ToDouble(InternalObject.GetType().InvokeMember("Height", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Height", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }
    }
}
