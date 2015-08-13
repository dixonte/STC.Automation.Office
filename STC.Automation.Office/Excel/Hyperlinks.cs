using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using System.IO;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Wraps an Excel.Hyperlinks object
    /// </summary>
    [WrapsCOM("Excel.Hyperlinks", "00024430-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class Hyperlinks : ComWrapper
    {
        internal Hyperlinks(object interiorObj)
            : base(interiorObj)
        {
        }

        /// <summary>
        /// Adds a hyperlink to the specified range or shape. The returned hyperlink must be manually disposed.
        /// </summary>
        /// <param name="anchor">The anchor for the hyperlink. Can be either a Range or Shape object.</param>
        /// <param name="address">The address of the hyperlink.</param>
        /// <param name="subAddress">The subaddress of the hyperlink.</param>
        /// <param name="screenTip">The screen tip to be displayed when the mouse pointer is paused over the hyperlink.</param>
        /// <param name="textToDisplay">The text to be displayed for the hyperlink.</param>
        /// <returns>A reference to the new Hyperlink</returns>
        public Hyperlink Add(Range anchor, string address, string subAddress = null, string screenTip = null, string textToDisplay = null)
        {
            return Add(anchor.InternalObject, address, subAddress, screenTip, textToDisplay);
        }

        public Hyperlink Add(Shape anchor, string address, string subAddress = null, string screenTip = null, string textToDisplay = null)
        {
            return Add(anchor.InternalObject, address, subAddress, screenTip, textToDisplay);
        }

        private Hyperlink Add(object anchor, string address, string subAddress = null, string screenTip = null, string textToDisplay = null)
        {
            return new Hyperlink(InternalObject.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, ComArguments.Prepare(anchor, address, subAddress, screenTip, textToDisplay)));
        }

        /// <summary>
        /// Index the hyperlinks
        /// </summary>
        /// <param name="id">hyperlink index</param>
        /// <returns></returns>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public Workbook this[int id]
        {
            get
            {
                return new Workbook(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, ComArguments.Prepare(id)));
            }
        }
    }
}
