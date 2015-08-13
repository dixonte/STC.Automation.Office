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
    /// Represents the interior of an object.
    /// </summary>
    [WrapsCOM("Excel.Axis", "00020848-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class Axis : ComWrapper
    {


        internal Axis(object interiorObj)
            : base(interiorObj)
        {
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
        /// Returns a TickLabels object that represents the tick-mark labels for the specified axis. Read-only.
        /// </summary>
        public TickLabels TickLabels
        {
            get
            {
                return new TickLabels(InternalObject.GetType().InvokeMember("TickLabels", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            
        }

        /// <summary>
        /// Returns an AxisTitle object that represents the title of the specified axis. Read-only.
        /// </summary>
        public AxisTitle AxisTitle
        {
            get
            {
                return new AxisTitle(InternalObject.GetType().InvokeMember("AxisTitle", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }


        }

        /// <summary>
        /// Returns or sets the minimum value on the value axis. Read/write Double.
        /// </summary>
        public double MinimumScale
        {
            get
            {
                return (double)InternalObject.GetType().InvokeMember("MinimumScale", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("MinimumScale", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns or sets the maximum value on the value axis. Read/write Double.
        /// </summary>
        public double MaximumScale
        {
            get
            {
                return (double)InternalObject.GetType().InvokeMember("MaximumScale", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("MaximumScale", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns or sets the major units for the value axis. Read/write Double.
        /// </summary>
        public double MajorUnit
        {
            get
            {
                return (double)InternalObject.GetType().InvokeMember("MajorUnit", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("MajorUnit", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }





    }
}
