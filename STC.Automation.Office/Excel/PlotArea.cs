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
    /// Represents the plot area of a chart.
    /// This is the area where your chart data is plotted. The plot area on a 2-D chart contains the data markers, gridlines, data labels, trendlines, and optional chart items placed in the chart area. The plot area on a 3-D chart contains all the above items plus the walls, floor, axes, axis titles, and tick-mark labels in the chart.
    /// The plot area is surrounded by the chart area. The chart area on a 2-D chart contains the axes, the chart title, the axis titles, and the legend. The chart area on a 3-D chart contains the chart title and the legend. For information about formatting the chart area, see the ChartArea object.
    /// </summary>
    [WrapsCOM("Excel.PlotArea", "000208CB-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class PlotArea : ComWrapper
    {
        private Border _border = null;

        internal PlotArea(object interiorObj)
            : base(interiorObj)
        {
        }


        /// <summary>
        /// Gets or sets the Interior of the range. (this is not automatically disposed)
        /// </summary>
        public Interior Interior
        {
            get
            {
                return new Interior(InternalObject.GetType().InvokeMember("Interior", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Interior", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Border of the plot area. This object is internally cached and does not need to be disposed.
        /// </summary>
        public Border Border
        {
            get
            {
                if (_border != null)
                    _border = new Border(InternalObject.GetType().InvokeMember("Border", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _border;
            }

        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {

                if (_border != null)
                {
                    _border.Dispose();
                    _border = null;
                }
            }

            base.Dispose(true);
        }

        #endregion




    }
}
