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
    /// Represents the chart area of a chart. 
    /// The chart area includes everything, including the plot area. However, the plot area has its own fill, so filling the plot area does not fill the chart area.
    ///For information about formatting the plot area, see PlotArea Object.
    ///Use the ChartArea property to return the ChartArea object. 
    /// </summary>
    [WrapsCOM("Excel.ChartArea", "000208CC-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class ChartArea : ComWrapper
    {
        private Font _font;

        internal ChartArea(object interiorObj)
            : base(interiorObj)
        {
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
