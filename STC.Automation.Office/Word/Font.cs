using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System.Drawing;
using STC.Automation.Office.Word.Enums;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Contains font attributes (such as font name, font size and color) for an object.
    /// </summary>
    [WrapsCOM("Word._Font", "00020952-0000-0000-C000-000000000046")]
    public class Font : ComWrapper
    {
        internal Font(object fontObj)
            : base(fontObj)
        {
        }

        /// <summary>
        /// Gets or sets the AllCaps status of the font. True if the font is formatted as AllCaps. Returns null if there is a mixture. Set null to toggle.
        /// Note that toggling a mixture will set the whole selection AllCaps, rather than behaving as might be expected.
        /// This setting is mutually exclusive with SmallCaps.
        /// </summary>
        public bool? AllCaps
        {
            get
            {
                return GetTernaryValue(Convert.ToInt64(InternalObject.GetType().InvokeMember("AllCaps", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null)));
            }

            set
            {
                InternalObject.GetType().InvokeMember("AllCaps", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { SetTernaryValue(value) });
            }
        }

        /// <summary>
        /// Gets or sets the bold status of the font. True if the font is formatted as bold. Returns null if there is a mixture. Set null to toggle.
        /// Note that toggling a mixture will set the whole selection bold, rather than behaving as might be expected.
        /// </summary>
        public bool? Bold
        {
            get
            {
                return GetTernaryValue(Convert.ToInt64(InternalObject.GetType().InvokeMember("Bold", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null)));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Bold", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { SetTernaryValue(value) });
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
        /// Gets or sets the italic status of the font. True if the font is formatted as italic. Returns null if there is a mixture. Set null to toggle.
        /// Note that toggling a mixture will set the whole selection italic, rather than behaving as might be expected.
        /// </summary>
        public bool? Italic
        {
            get
            {
                return GetTernaryValue(Convert.ToInt64(InternalObject.GetType().InvokeMember("Italic", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null)));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Italic", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { SetTernaryValue(value) });
            }
        }

        /// <summary>
        /// Gets or sets the name of the font.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the SmallCaps status of the font. True if the font is formatted as SmallCaps. Returns null if there is a mixture. Set null to toggle.
        /// Note that toggling a mixture will set the whole selection SmallCaps, rather than behaving as might be expected.
        /// This setting is mutually exclusive with AllCaps.
        /// </summary>
        public bool? SmallCaps
        {
            get
            {
                return GetTernaryValue(Convert.ToInt64(InternalObject.GetType().InvokeMember("SmallCaps", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null)));
            }

            set
            {
                InternalObject.GetType().InvokeMember("SmallCaps", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { SetTernaryValue(value) });
            }
        }

        /// <summary>
        /// Gets or sets the Subscript status of the font. True if the font is formatted as Subscript. Returns null if there is a mixture. Set null to toggle.
        /// Note that toggling a mixture will set the whole selection Subscript, rather than behaving as might be expected.
        /// This setting is mutually exclusive with Superscript.
        /// </summary>
        public bool? Subscript
        {
            get
            {
                return GetTernaryValue(Convert.ToInt64(InternalObject.GetType().InvokeMember("Subscript", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null)));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Subscript", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { SetTernaryValue(value) });
            }
        }

        /// <summary>
        /// Gets or sets the Superscript status of the font. True if the font is formatted as Superscript. Returns null if there is a mixture. Set null to toggle.
        /// Note that toggling a mixture will set the whole selection Superscript, rather than behaving as might be expected.
        /// This setting is mutually exclusive with Subscript.
        /// </summary>
        public bool? Superscript
        {
            get
            {
                return GetTernaryValue(Convert.ToInt64(InternalObject.GetType().InvokeMember("Superscript", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null)));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Superscript", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { SetTernaryValue(value) });
            }
        }

        #region Private Methods

        private static bool? GetTernaryValue(long value)
        {
            if (value == (long)Constants.Undefined)
            {
                return null;
            }
            else
            {
                return (value != 0);
            }
        }

        private static object SetTernaryValue(bool? value)
        {
            return (value != null) ? (object)value : (object)Constants.Toggle;
        }

        #endregion
    }
}
