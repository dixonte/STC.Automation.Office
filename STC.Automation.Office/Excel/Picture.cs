using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Wraps an Excel.Picture object. This class is hidden in the Excel OLE model and should not be relied upon.
    /// </summary>
    [Obsolete("This class is hidden in the Excel OLE model and should not be relied upon.")]
    [WrapsCOM("Excel.Picture", "000208A6-0000-0000-C000-000000000046")]
    public class Picture : ComWrapper
    {
        internal Picture(object pictureObj)
            : base(pictureObj)
        {
        }

        /// <summary>
        /// Gets or sets the name of this picture.
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
