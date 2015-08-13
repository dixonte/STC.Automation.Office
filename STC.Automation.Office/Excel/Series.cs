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
    /// Represents a series in a chart.
    /// </summary>
    [WrapsCOM("Excel.Series", "0002086B-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class Series : ComWrapper
    {
        internal Series(object interiorObj)
            : base(interiorObj)
        {
        }


        /// <summary>
        /// Returns or sets a String value representing the name of the series.
        /// </summary>
        public string Name
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null).ToString();

            }

            set
            {
                InternalObject.GetType().InvokeMember("Name", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets a Borderobject for the series
        /// </summary>
        public Border Border
        {
            get
            {
                return new Border(InternalObject.GetType().InvokeMember("Border", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Returns or sets an array of x values for a chart series. 
        /// </summary>
        public string XValues
        {
            get
            {
                return InternalObject.GetType().InvokeMember("XValues", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null).ToString();
            }

            set
            {
                InternalObject.GetType().InvokeMember("XValues", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns or sets a Variant value that represents a collection of all the values in the series. 
        /// </summary>
        public string Values
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Values", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null).ToString();
            }

            set
            {
                InternalObject.GetType().InvokeMember("Values", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets a Borderobject for the series
        /// </summary>
        public MarkerStyles MarkerStyle
        {
            get
            {
                return (MarkerStyles)InternalObject.GetType().InvokeMember("MarkerStyle", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("MarkerStyle", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns an XlAxisGroup value that represents the type of axis group. Read/write.
        /// </summary>
        public AxisGroup AxisGroup
        {
            get
            {
                return (AxisGroup)InternalObject.GetType().InvokeMember("AxisGroup", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("AxisGroup", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { (int)value });
            }
        }

        /// <summary>
        /// Returns or sets the chart type. Read/write XlChartType.
        /// </summary>
        public ChartType ChartType
        {
            get
            {
                return (ChartType)InternalObject.GetType().InvokeMember("ChartType", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("ChartType", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { (int)value });
            }
        }

    }
}
