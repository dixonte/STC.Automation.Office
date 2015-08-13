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
    [WrapsCOM("Excel.SeriesCollection", "0002086C-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class SeriesCollection : ComWrapper
    {


        internal SeriesCollection(object interiorObj)
            : base(interiorObj)
        {
        }

        /// <summary>
        /// Adds one or more new series to the SeriesCollection collection.
        /// </summary>
        /// <param name="Source">The new data as a Range object.</param>
        /// <param name="rowcol">Specifies whether the new values are in the rows or columns of the specified range.</param>
        /// <param name="serieslabels">True if the first row or column contains the name of the data series. False if the first row or column contains the first data point of the series. If this argument is omitted, Microsoft Excel attempts to determine the location of the series name from the contents of the first row or column.</param>
        /// <param name="categoryLabels">True if the first row or column contains the name of the category labels. False if the first row or column contains the first data point of the series. If this argument is omitted, Microsoft Excel attempts to determine the location of the category label from the contents of the first row or column.</param>
        /// <param name="replace">If CategoryLabels is True and Replace is True, the specified categories replace the categories that currently exist for the series. If Replace is False, the existing categories will not be replaced. The default value is False.</param>
        /// <returns></returns>
        public Series Add(Range Source, RowCol? rowcol = null, bool? serieslabels = null, bool? categoryLabels = null, bool? replace = null)
        {
            return new Series(InternalObject.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { Source, rowcol, serieslabels, categoryLabels, replace }));
        }

        /// <summary>
        /// Creates a new series. Returns a Series object that represents the new series.
        /// </summary>
        /// <returns>a series object</returns>
        public Series NewSeries()
        {
            return new Series(InternalObject.GetType().InvokeMember("NewSeries", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null));
        }




        /// <summary>
        /// Index the Series collection to get a series
        /// </summary>
        /// <param name="key">Series name</param>
        /// <returns>A Series</returns>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public Series this[string key]
        {
            get
            {
                try
                {
                    return new Series(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { key }));
                }
                catch (Exception ex)
                {
                    throw new IndexOutOfRangeException(String.Concat("Could not find series '", key, "'."), ex);
                }
            }
        }
        

        /// <summary>
        /// Index the Series collection to get a series
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public Series this[int key]
        {
            get
            {
                try
                {
                    return new Series(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { key }));
                }
                catch (Exception ex)
                {
                    throw new IndexOutOfRangeException(String.Concat("Could not find series '", key, "'."), ex);
                }
            }
        }


        /// <summary>
        /// Returns an integer value that represents the number of objects in the collection.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public int Count
        {
            get
            {
                return (int)InternalObject.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }
        }

    }
}
