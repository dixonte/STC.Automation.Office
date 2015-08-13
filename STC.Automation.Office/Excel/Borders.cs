using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using STC.Automation.Office.Excel.Enums;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// A collection of four Border objects that represent the four borders of a Range or Style object.
    /// </summary>
    [WrapsCOM("Excel.Borders", "00020855-0000-0000-C000-000000000046")]
    public class Borders : ComWrapper
    {
        private Dictionary<BordersIndex, Border> _this = new Dictionary<BordersIndex,Border>();

        internal Borders(object bordersObj)
            : base(bordersObj)
        {
        }

        /// <summary>
        /// Gets a Border object that represents one of the borders of either a range of cells or a style. This object is internally cached and does not need to be manually disposed.
        /// </summary>
        /// <param name="idx">A variable that represents a Borders object.</param>
        /// <returns>Border</returns>
        public Border this[BordersIndex idx]
        {
            get
            {
                if (!_this.ContainsKey(idx))
                {
                    _this.Add(idx, new Border(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { idx })));
                }

                return _this[idx];
            }
        }

        internal override void Dispose(bool disposing)
        {
            foreach (Border border in _this.Values)
            {
                border.Dispose();
            }
            _this.Clear();

            base.Dispose(disposing);
        }
    }
}
