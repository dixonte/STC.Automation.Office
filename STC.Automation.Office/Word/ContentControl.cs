using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Represents a single ContentControl in a document.
    /// </summary>
    [WrapsCOM("Word.ContentControl", "EE95AFE3-3026-4172-B078-0E79DAB5CC3D")]
    public class ContentControl : ComWrapper
    {
        internal ContentControl(object contentControlObj)
            : base(contentControlObj)
        {
        }

        /// <summary>
        /// Deletes the specified ContentControl.
        /// </summary>
        /// <param name="deleteContents">Specifies whether to delete the contents of the content control. True removes both the content control and its contents. False removes the control but leaves the contents of the content control in the active document. The default value is False.</param>
        public void Delete(bool? deleteContents)
        {
            List<object> args = new List<object>();
            if (deleteContents.HasValue)
                args.Add(deleteContents.Value);

            InternalObject.GetType().InvokeMember("Delete", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, args.ToArray());
        }

        /// <summary>
        /// Returns a Range that represents the contents of the content control in the active document.
        /// This Range object is NOT internally cached and must be manually disposed.
        /// </summary>
        public Range Range
        {
            get
            {
                return new Range(InternalObject.GetType().InvokeMember("Range", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Gets the Title of the ContentControl.
        /// </summary>
        public string Title
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Title", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null) as string;
            }
        }


        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
            }

            base.Dispose(true);
        }

        #endregion
    }
}
