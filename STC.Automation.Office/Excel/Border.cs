using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using STC.Automation.Office.Excel.Enums;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Represents the border of an object.
    /// </summary>
    [WrapsCOM("Excel.Border", "00020854-0000-0000-C000-000000000046")]
    public class Border : ComWrapper
    {
        internal Border(object borderObj)
            : base(borderObj)
        {
        }

        /// <summary>
        /// Gets or sets the line style for the border.
        /// </summary>
        public LineStyle LineStyle
        {
            get
            {
                return (LineStyle)InternalObject.GetType().InvokeMember("LineStyle", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("LineStyle", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the 24-bit color for the specified Border object.
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
        /// Gets or sets an integer value that represents the color of the border. The color is specified as an index value into the current color palette, or as one of the ColorIndex enum values.
        /// </summary>
        public int ColorIndex
        {
            get
            {
                return Convert.ToInt32(InternalObject.GetType().InvokeMember("ColorIndex", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("ColorIndex", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets a Single that lightens or darkens a color.
        /// </summary>
        /// <remarks>
        /// You can enter a number from -1 (darkest) to 1 (lightest) for the TintAndShade  property. Zero (0) is neutral.
        /// Attempting to set this property to a value less than -1 or more than 1 results in a run-time error: "The specified value is out of range." This property works for both theme colors and nontheme colors.
        /// </remarks>
        public Single TintAndShade
        {
            get
            {
                return Convert.ToSingle(InternalObject.GetType().InvokeMember("TintAndShade", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("TintAndShade", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets a BorderWeight enum value that represents the weight of the border.
        /// </summary>
        public BorderWeight Weight
        {
            get
            {
                return (BorderWeight)InternalObject.GetType().InvokeMember("Weight", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("Weight", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }
    }
}
