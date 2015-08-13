using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System.Drawing;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Contains the font attributes (font name, font size, color, and so on) for an object.
    /// </summary>
    [WrapsCOM("Excel.Font", "0002084D-0000-0000-C000-000000000046")]
    public class Font : ComWrapper
    {
        internal Font(object fontObj)
            : base(fontObj)
        {
        }

        /// <summary>
        /// Gets or sets a boolean value as to whether the font is bold or not.
        /// </summary>
        public bool Bold
        {
            get
            {
                return Convert.ToBoolean(InternalObject.GetType().InvokeMember("Bold", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Bold", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the 24-bit color for the specified Font object.
        /// </summary>
        public Color Color
        {
            get
            {
                return ColorTranslator.FromOle(Convert.ToInt32(InternalObject.GetType().InvokeMember("Color", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null)));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Color", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { ColorTranslator.ToOle(value) });
            }
        }

        /// <summary>
        /// Gets or sets a boolean value as to whether the font is italic or not.
        /// </summary>
        public bool Italic
        {
            get
            {
                return Convert.ToBoolean(InternalObject.GetType().InvokeMember("Italic", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Italic", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the name of the font.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public string Name
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null) as string;
            }

            set
            {
                InternalObject.GetType().InvokeMember("Name", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the font size, in points.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public Single Size
        {
            get
            {
                return Convert.ToSingle(InternalObject.GetType().InvokeMember("Size", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Size", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }


    }
}
