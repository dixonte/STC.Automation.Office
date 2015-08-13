using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using STC.Automation.Office.Word.Enums;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Represents the criteria for a find operation. 
    /// </summary>
    [WrapsCOM("Word.Find", "000209B0-0000-0000-C000-000000000046")]
    public class Find : ComWrapper
    {
        private Replacement _replacement;

        internal Find(object findObj)
            : base(findObj)
        {
        }

        /// <summary>
        /// Removes text and paragraph formatting from the text specified in a find or replace operation.
        /// </summary>
        public void ClearFormatting()
        {
            InternalObject.GetType().InvokeMember("ClearFormatting", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        // TODO: Add some overloads to this.
        /// <summary>
        /// Runs the specified find operation. 
        /// </summary>
        /// <param name="replace">Specifies how many replacements are to be made: one, all, or none. Can be any Replace constant.</param>
        /// <returns>True if the find operation is successful.</returns>
        public bool Execute(Replace replace)
        {
            return Convert.ToBoolean(InternalObject.GetType().InvokeMember("Execute", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject,
                new object[] { replace },
                null,
                System.Threading.Thread.CurrentThread.CurrentCulture,
                new string[] { "Replace" }));
        }

        /// <summary>
        /// Returns a Replacement object that contains the criteria for a replace operation.
        /// This object is internally cached and does not need to be manually disposed.
        /// </summary>
        public Replacement Replacement
        {
            get
            {
                if (_replacement == null)
                    _replacement = new Replacement(InternalObject.GetType().InvokeMember("Replacement", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _replacement;
            }
        }

        /// <summary>
        /// Gets or sets the text to find.
        /// </summary>
        public string Text
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Text", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null) as string;
            }

            set
            {
                InternalObject.GetType().InvokeMember("Text", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_replacement != null)
                {
                    _replacement.Dispose();
                    _replacement = null;
                }
            }

            base.Dispose(true);
        }

        #endregion
    }
}
