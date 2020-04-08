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
    /// Represents an embedded chart on a worksheet.
    /// The ChartObject object acts as a container for a Chart object. Properties and methods for the ChartObject object control the appearance and size of the embedded chart on the worksheet. 
    /// The ChartObject object is a member of the ChartObjects collection. The ChartObjects collection contains all the embedded charts on a single sheet.
    /// </summary>
    [WrapsCOM("Excel.ChartObject", "000208CF-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class ChartObject : ComWrapper
    {
        private Chart _chart;

        internal ChartObject(object interiorObj)
            : base(interiorObj)
        {
        }

        /// <summary>
        /// Makes the current chart the active chart.
        /// </summary>
        public ChartObject Activate()
        {
            return new ChartObject(InternalObject.GetType().InvokeMember("Activate", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null));
        }

        /// <summary>
        /// Returns a Chart object that represents the chart contained in the object. Read-only.
        /// </summary>
        /// <returns>A Chart object</returns>
        public Chart Chart
        {
            get
            {
                if (_chart == null)
                {
                    _chart = new Chart(InternalObject.GetType().InvokeMember("Chart", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _chart;
            }
        }

        internal override void Dispose(bool disposing)
        {
            if (_chart != null)
            {
                _chart.Dispose();
                _chart = null;
            }

            base.Dispose(disposing);
        }

    }
}
