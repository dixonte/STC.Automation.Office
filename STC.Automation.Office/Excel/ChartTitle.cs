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
    /// Represents the title of a chart
    /// </summary>
    [WrapsCOM("Excel.ChartTitle", "00020849-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class ChartTitle : ComWrapper
    {
        private Font _font = null;

        internal ChartTitle(object interiorObj)
            : base(interiorObj)
        {
        }


        /// <summary>
        /// Returns or sets the text for the specified object. Read/write String.
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
        /// Font object
        /// </summary>
        public Font Font
        {
            get
            {
                if (_font != null)
                    _font = new Font(InternalObject.GetType().InvokeMember("Font", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _font;
            }


        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
               
                if (_font != null)
                {
                    _font.Dispose();
                    _font = null;
                }
            }

            base.Dispose(true);
        }

        #endregion


    }
}
