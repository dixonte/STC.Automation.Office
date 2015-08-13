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
    /// Represents the legend in a chart. Each chart can have only one legend. 
    /// The Legend object contains one or more LegendEntry objects; each LegendEntry object contains a LegendKey object.
    /// The chart legend isn't visible unless the HasLegend property is True. If this property is False, properties and methods of the Legend object will fail.
    /// </summary>
    [WrapsCOM("Excel.Legend", "000208CD-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class Legend : ComWrapper
    {
        internal Legend(object interiorObj)
            : base(interiorObj)
        {
        }

        /// <summary>
        /// Returns or sets a LegendPosition value that represents the position of the legend on the chart.
        /// </summary>
        public LegendPosition Position
        {
            get
            {
                return (LegendPosition)InternalObject.GetType().InvokeMember("Position", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("Position", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }



        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // dispose internal objects
            }

            base.Dispose(true);
        }

        #endregion

    }
}
