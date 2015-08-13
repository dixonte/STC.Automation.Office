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
    /// The chart can be either an embedded chart (contained in a ChartObject object) or a separate chart sheet.
    /// </summary>
    [WrapsCOM("Excel.Chart", "000208D6-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class Chart : Sheet
    {
        private SeriesCollection _seriescollection = null;
        private Axes _axes = null;
        private ChartTitle _chartTitle = null;
        private ChartArea _chartArea = null;
        private PlotArea _plotArea = null;
        private Legend _legend = null;

        internal Chart(object interiorObj)
            : base(interiorObj)
        {
        }


        /// <summary>
        /// Returns or sets the chart type.
        /// </summary>
        public ChartType ChartType
        {
            get
            {
                return (ChartType)InternalObject.GetType().InvokeMember("ChartType", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }
            set
            {
                InternalObject.GetType().InvokeMember("ChartType", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns a ChartArea object that represents the complete chart area for the chart. Read-only.
        /// </summary>
        public ChartArea ChartArea
        {
            get
            {
                if (_chartArea != null)
                    _chartArea = new ChartArea(InternalObject.GetType().InvokeMember("ChartArea", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _chartArea;
            }
        }

        /// <summary>
        /// Returns a PlotArea object that represents the plot area of a chart. Read-only. This object is internally cached and does not need to be disposed.
        /// </summary>
        public PlotArea PlotArea
        {
            get
            {
                if (_plotArea != null)
                    _plotArea = new PlotArea(InternalObject.GetType().InvokeMember("PlotArea", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _plotArea;
            }

        }

        /// <summary>
        /// Returns an object that represents either a single series (a Series object) or a collection of all the series (a SeriesCollection collection) in the chart or chart group.
        /// This object is cached internally and does not need to be disposed.
        /// </summary>
        public SeriesCollection SeriesCollection
        {
            
            get
            {
                if (_seriescollection == null)
                {
                    _seriescollection = new SeriesCollection(InternalObject.GetType().InvokeMember("SeriesCollection", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _seriescollection;
            }
      
        }

        /// <summary>
        /// Returns a ChartTitle object that represents the title of the specified chart. Read-only.
        /// This class is internally cached and does not need to be disposed.
        /// </summary>
        public ChartTitle ChartTitle
        {

            get
            {
                if (_chartTitle == null)
                    _chartTitle = new ChartTitle(InternalObject.GetType().InvokeMember("ChartTitle", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _chartTitle;
            }

        }

        /// <summary>
        /// The Axes for the chart
        /// </summary>
        public Axes Axes
        {

            get
            {
                if (_axes == null)
                {
                    _axes = new Axes(InternalObject.GetType().InvokeMember("Axes", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _axes;
            }

        }

        /// <summary>
        /// True if the axis or chart has a visible title. Read/write Boolean.
        /// </summary>
        public bool HasTitle
        {
            get
            {
                return (bool)InternalObject.GetType().InvokeMember("HasTitle", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("HasTitle", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// True if the chart has a legend. Read/write Boolean.
        /// </summary>
        public bool HasLegend
        {
            get
            {
                return (bool)InternalObject.GetType().InvokeMember("HasLegend", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("HasLegend", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns a Legend object that represents the legend for the chart. Read-only.
        /// </summary>
        public Legend Legend
        {
            get
            {
                if (_legend == null)
                    _legend = new Legend(InternalObject.GetType().InvokeMember("Legend", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _legend;
            }

        }

        /// <summary>
        /// Sets the source data range for the chart.
        /// </summary>
        /// <param name="source">The range that contains the source data.</param>
        /// <param name="plotBy">Specifies the way the data is to be plotted.</param>
        public void SetSourceData(Range source, RowCol plotBy)
        {
            var missing = System.Reflection.Missing.Value;

            InternalObject.GetType().InvokeMember("SetSourceData", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { (source != null ? source.InternalObject : missing), (int)plotBy });
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_seriescollection != null)
                {
                    _seriescollection.Dispose();
                    _seriescollection = null;
                }

                if (_axes != null)
                {
                    _axes.Dispose();
                    _axes = null;
                }

                if (_chartTitle != null)
                {
                    _chartTitle.Dispose();
                    _chartTitle = null;
                }

                if (_chartArea != null)
                {
                    _chartArea.Dispose();
                    _chartArea = null;
                }

                if (_plotArea != null)
                {
                    _plotArea.Dispose();
                    _plotArea = null;
                }

                if (_legend != null)
                {
                    _legend.Dispose();
                    _legend = null;
                }
            }

            base.Dispose(true);
        }

        #endregion
    }
}
