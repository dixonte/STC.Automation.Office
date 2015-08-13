using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using STC.Automation.Office.Word.Enums;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Represents all the formatting for a paragraph.
    /// </summary>
    [WrapsCOM("Word._ParagraphFormat", "00020953-0000-0000-C000-000000000046")]
    public class ParagraphFormat : ComWrapper
    {
        /// <summary>
        /// Creates a new wrapped ParagraphFormat object.
        /// </summary>
        public ParagraphFormat()
            : base()
        {
        }

        internal ParagraphFormat(object paragraphFormatObj)
            : base(paragraphFormatObj)
        {
        }

        /// <summary>
        /// Gets or sets a ParagraphAlignment constant that represents the alignment for the specified paragraphs.
        /// </summary>
        public ParagraphAlignment Alignment
        {
            get
            {
                return (ParagraphAlignment)Convert.ToInt32(InternalObject.GetType().InvokeMember("Alignment", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Alignment", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets a Single that represents the left indent value (in points) for the specified paragraph formatting.
        /// </summary>
        public Single LeftIndent
        {
            get
            {
                return Convert.ToSingle(InternalObject.GetType().InvokeMember("LeftIndent", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("LeftIndent", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets a Single that represents the right indent value (in points) for the specified paragraph formatting.
        /// </summary>
        public Single RightIndent
        {
            get
            {
                return Convert.ToSingle(InternalObject.GetType().InvokeMember("RightIndent", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("RightIndent", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// True if Microsoft Word automatically sets the amount of spacing after the specified paragraphs.
        /// Returns null if the SpaceAfterAuto property is set to True for only some of the specified paragraphs.
        /// Cannot be set null.
        /// </summary>
        public bool? SpaceAfterAuto
        {
            get
            {
                var val = Convert.ToInt64(InternalObject.GetType().InvokeMember("SpaceAfterAuto", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                if (val == (long)Constants.Undefined)
                {
                    return null;
                }
                else
                {
                    return (val != 0);
                }
            }

            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException("SpaceAfterAuto can only be set to True or False.");
                }

                InternalObject.GetType().InvokeMember("SpaceAfterAuto", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// True if Microsoft Word automatically sets the amount of spacing before the specified paragraphs.
        /// Returns null if the SpaceBeforeAuto property is set to True for only some of the specified paragraphs.
        /// Cannot be set null.
        /// </summary>
        public bool? SpaceBeforeAuto
        {
            get
            {
                var val = Convert.ToInt64(InternalObject.GetType().InvokeMember("SpaceBeforeAuto", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                if (val == (long)Constants.Undefined)
                {
                    return null;
                }
                else
                {
                    return (val != 0);
                }
            }

            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException("SpaceBeforeAuto can only be set to True or False.");
                }

                InternalObject.GetType().InvokeMember("SpaceBeforeAuto", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }
    }
}
