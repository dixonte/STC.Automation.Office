using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using STC.Automation.Office.Core.Enums;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Represents the color of a one-color object, the foreground or background color of an object with a gradient or patterned fill, or the pointer color.
    /// </summary>
    [WrapsCOM("Excel.ColorFormat", "000C0312-0000-0000-C000-000000000046")]
    public class ColorFormat: ComWrapper
    {
        internal ColorFormat(object interiorObj)
               : base(interiorObj)
        {

        }

        /// <summary>
        /// Returns or sets a color that is mapped to the theme color scheme. Read/write MsoThemeColorIndex.
        /// </summary>
        public ThemeColorIndex ObjectThemeColor
        {
            get
            {
                return (ThemeColorIndex)(InternalObject.GetType().InvokeMember("ObjectThemeColor", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("ObjectThemeColor", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Mimics the VBA RGB Function which combined the red, green and blue components into a single numeric value
        /// </summary>
        /// <param name="red">The red component (0 - 255)</param>
        /// <param name="green">The green component (0 - 255)</param>
        /// <param name="blue">The green component (0-255)</param>
        /// <returns></returns>
        public static long RGBToLong(int red, int green, int blue)
        {
            return red + (green * 256) + (blue * 65536);
        }

        /// <summary>
        /// Gets or sets the red-green-blue value of the specified color.
        /// </summary>
        public long RGB
        {
            get
            {
                return (long)(InternalObject.GetType().InvokeMember("RGB", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("RGB", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }
    }
}
