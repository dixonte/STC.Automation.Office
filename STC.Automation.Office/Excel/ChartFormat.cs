using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Provides access to the Office Art formatting for chart elements.
    /// </summary>
    [WrapsCOM("Excel.Series", "000244B2-0000-0000-C000-000000000046")]
    public class ChartFormat : ComWrapper
    {
        private FillFormat _fillFormat;

        internal ChartFormat(object interiorObj)
               : base(interiorObj)
        {
        }

        public FillFormat Fill
        {
            get

            {
                if (_fillFormat == null)
                {
                    _fillFormat = new FillFormat(InternalObject.GetType().InvokeMember("Fill", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _fillFormat;
            }
        }

        internal override void Dispose(bool disposing)
        {
            if (_fillFormat != null)
            {
                _fillFormat.Dispose();
                _fillFormat = null;
            }

            base.Dispose(disposing);
        }
    }
}
