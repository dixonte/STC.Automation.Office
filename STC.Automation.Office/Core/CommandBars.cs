using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Word.Enums;

namespace STC.Automation.Office.Core
{
    /// <summary>
    /// Wraps an Core.CommandBars COM object
    /// </summary>
    [WrapsCOM("Core.CommandBars", "000C0302-0000-0000-C000-000000000046")]
    public class CommandBars : ComWrapper, IEnumerable<CommandBar>
    {
        /// <summary>
        /// Wraps the given COM object as a CommandBars.
        /// </summary>
        /// <param name="commandBarsObj"></param>
        public CommandBars(object commandBarsObj)
            : base(commandBarsObj)
        {
        }

        public CommandBar this[int index]
        {
            get
            {
                return new CommandBar(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { index }));
            }
        }

        public CommandBar this[string key]
        {
            get
            {
                return new CommandBar(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { key }));
            }
        }

        public int Count
        {
            get
            {
                return (int)InternalObject.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
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

        #region IEnumerable<CommandBar> Members

        public IEnumerator<CommandBar> GetEnumerator()
        {
            for (int x = 1; x <= Count; x++)
                yield return this[x];
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            for (int x = 1; x <= Count; x++)
                yield return this[x];
        }

        #endregion
    }
}
