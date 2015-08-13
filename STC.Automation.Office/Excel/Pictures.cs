using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Common;
using System.IO;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Wraps an Excel.Pictures collection object.
    /// </summary>
    [WrapsCOM("Excel.Pictures", "000208A7-0000-0000-C000-000000000046")]
    public class Pictures : ComWrapper
    {
        internal Pictures(object picturesObj)
            : base(picturesObj)
        {
        }

        /// <summary>
        /// Inserts an image into the active cell of the worksheet to which this Pictures object belongs. The returned Picture object must be manually disposed.
        /// </summary>
        /// <param name="filename">Full path to the image file</param>
        /// <returns>A Picture object</returns>
        [Obsolete("This method is hidden in the Excel OLE model and should not be relied upon. Use Shapes.AddPicture instead.")]
        public Picture Insert(string filename)
        {
            if (!File.Exists(filename))
                throw new FileNotFoundException("Could not find file " + filename);

            return new Picture(InternalObject.GetType().InvokeMember("Insert", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { filename }));
        }
    }
}
