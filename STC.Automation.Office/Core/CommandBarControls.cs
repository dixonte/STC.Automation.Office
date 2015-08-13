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
    /// Wraps an Core.CommandBarControls COM object
    /// </summary>
    [WrapsCOM("Core.CommandBarControls", "000C0306-0000-0000-C000-000000000046")]
    public class CommandBarControls : ComWrapper, IEnumerable<CommandBarControl>
    {
        /// <summary>
        /// Wraps the given COM object as a CommandBarControls.
        /// </summary>
        /// <param name="commandBarControlsObj"></param>
        public CommandBarControls(object commandBarControlsObj)
            : base(commandBarControlsObj)
        {
        }

        public CommandBarControl this[int index]
        {
            get
            {
                return WrapCorrectType(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { index }));
            }
        }

        public CommandBarControl this[string key]
        {
            get
            {
                return WrapCorrectType(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { key }));
            }
        }

        private CommandBarControl WrapCorrectType(object toWrap)
        {
            if (SupportsInterface(toWrap, new Guid(CommandBarButton.UUID)))
                return new CommandBarButton(toWrap);
            else if (SupportsInterface(toWrap, new Guid(CommandBarComboBox.UUID)))
                return new CommandBarComboBox(toWrap);
            else if (SupportsInterface(toWrap, new Guid(CommandBarPopup.UUID)))
                return new CommandBarPopup(toWrap);

            return new CommandBarControl(toWrap);
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

        #region IEnumerable<CommandBarControl> Members

        public IEnumerator<CommandBarControl> GetEnumerator()
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
