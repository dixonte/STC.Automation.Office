using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office
{
    /// <summary>
    /// In a series, the DataLabel object is a member of the DataLabels collection. The DataLabels collection contains a DataLabel object for each point. For a series without definable points (such as an area series), the DataLabels collection contains a single DataLabel object.
    /// </summary>
    [WrapsCOM("Excel.PivotCache", "000208B2-0000-0000-C000-000000000046")]
    public class DataLabel : ComWrapper
    {
        internal DataLabel(object intObj)
            : base(intObj)
        {
        }

        /// <summary>
        /// Returns or sets a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide. Read/write.
        /// </summary>
        public bool ShowValue
        {
            get
            {
                return (bool)(InternalObject.GetType().InvokeMember("ShowValue", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("ShowValue", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// True if the data label legend key is visible. Read/write Boolean.
        /// </summary>
        public bool ShowLegendKey
        {
            get
            {
                return (bool)(InternalObject.GetType().InvokeMember("ShowLegendKey", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("ShowLegendKey", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }
    }
}
