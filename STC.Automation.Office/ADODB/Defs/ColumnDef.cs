using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.ADODB.Defs
{
    /// <summary>
    /// Defines important details for ADODB.Field objects when created by ADODB.Recordset.FromDataTable.
    /// </summary>
    public class ColumnDef
    {
        /// <summary>
        /// Column name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Defined size of field.
        /// </summary>
        public long? DefinedSize { get; set; }
        /// <summary>
        /// Maximum number of digits used to represet values.
        /// </summary>
        public byte? Precision { get; set; }
        /// <summary>
        /// Number of decimal places to which numeric values will be resolved.
        /// </summary>
        public byte? NumericScale { get; set; }
    }
}
