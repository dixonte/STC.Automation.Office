using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using STC.Automation.Office.Core.Enums;
using STC.Automation.Office.Excel.Enums;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel
{
    [WrapsCOM("Excel.FillFormat", "000C0314-0000-0000-C000-000000000046")]
    public class FillFormat : ComWrapper
    {
        internal FillFormat(object interiorObj)
               : base(interiorObj)
        {
        }

        /// <summary>
        /// Returns or sets an MsoTriState value that determines whether the object is visible. Read/write.
        /// </summary>
        public TriState Visible
        {
            get
            {
                return (TriState)InternalObject.GetType().InvokeMember("Visible", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns or sets a ColorFormat object that represents the specified foreground fill or solid color.
        /// </summary>
        public ColorFormat ForeColor
        {
            get
            {
                return (new ColorFormat(InternalObject.GetType().InvokeMember("ForeColor", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null)));
            }

            set
            {
                InternalObject.GetType().InvokeMember("ForeColor", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }
    }
}
