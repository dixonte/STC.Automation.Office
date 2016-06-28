using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Data;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.ADODB
{
    /// <summary>
    /// Wraps an ADODB.Recordset object
    /// </summary>
    [WrapsCOM("ADODB.Recordset", "00000556-0000-0010-8000-00AA006D2EA4")]
    public class Recordset : ComWrapper
    {
        private Fields _fields;

        /// <summary>
        /// Creates a new, empty, disconnected Recordset.
        /// </summary>
        public Recordset()
            : base()
        {
        }

        internal Recordset(object recordsetObj)
            : base(recordsetObj)
        {
        }

        /// <summary>
        /// Builds a Recordset from a DataTable object.
        /// </summary>
        /// <param name="table">The source DataTable.</param>
        /// <param name="columnDefs">Optional. Provide size, precision, scale etc for columns in the DataTable.</param>
        /// <returns>A Recordset</returns>
        public static Recordset FromDataTable(DataTable table, IEnumerable<Defs.ColumnDef> columnDefs = null)
        {
            Recordset rs = new Recordset();

            var defs = new Dictionary<string, Defs.ColumnDef>();
            if (columnDefs != null)
            {
                foreach (var columnDef in columnDefs)
                {
                    if (string.IsNullOrEmpty(columnDef.Name))
                        throw new ArgumentException("columnDefs contains object with empty Name property.");

                    defs.Add(columnDef.Name, columnDef);
                }
            }

            // Add the columns to the recordset
            for (int x = 0; x < table.Columns.Count; x++)
            {
                DataColumn col = table.Columns[x];
                Enums.DataType rsDataType;
                long size;

                size = 0; // default used for data types which do not have size

                // figure out what the data type of the current column is
                if (col.DataType == typeof(Int16))
                {
                    rsDataType = STC.Automation.Office.ADODB.Enums.DataType.SmallInt;
                }
                else if  (col.DataType == typeof(Int32))
                {
                    rsDataType = STC.Automation.Office.ADODB.Enums.DataType.Integer;                         
                }
                else if (col.DataType == typeof(Int64))
                {
                    rsDataType = STC.Automation.Office.ADODB.Enums.DataType.BigInt;
                }
                else if (col.DataType == typeof(Boolean))
                {
                    rsDataType = STC.Automation.Office.ADODB.Enums.DataType.Boolean;
                }
                else if (col.DataType == typeof(DateTime))
                {
                    rsDataType = STC.Automation.Office.ADODB.Enums.DataType.Date;
                }
                else if (col.DataType == typeof(Decimal))
                {
                    rsDataType = STC.Automation.Office.ADODB.Enums.DataType.Decimal;
                }
                else if (col.DataType == typeof(Double))
                {
                    rsDataType = STC.Automation.Office.ADODB.Enums.DataType.Double;
                }
                else if (col.DataType == typeof(Guid))
                {
                    rsDataType = STC.Automation.Office.ADODB.Enums.DataType.GUID;
                }
                else if (col.DataType == typeof(String))
                {
                    rsDataType = STC.Automation.Office.ADODB.Enums.DataType.VarWChar;
                    size = 1000; // should be sufficient for most large strings.
                }
                else if (col.DataType == typeof(Single))
                {
                    rsDataType = STC.Automation.Office.ADODB.Enums.DataType.Single;
                }
                else if (col.DataType == typeof(byte))
                {
                    rsDataType = STC.Automation.Office.ADODB.Enums.DataType.TinyInt;
                }
                else
                {
                    throw new NotImplementedException("Unsupported datatype found in DataTable during converion to Recordset");
                }

                Defs.ColumnDef def = null;
                if (defs.ContainsKey(col.ColumnName))
                {
                    def = defs[col.ColumnName];
                    if (def.DefinedSize.HasValue)
                        size = def.DefinedSize.Value;
                }

                // create the column
                if (size == 0)
                {
                    // size is not relevant
                    rs.Fields.Append(col.Caption, rsDataType, STC.Automation.Office.ADODB.Enums.FieldAttribute.MayBeNull);
                }
                else
                {
                    // size is relevant
                    rs.Fields.Append(col.Caption, rsDataType, size, STC.Automation.Office.ADODB.Enums.FieldAttribute.MayBeNull);
                }

                if (def?.Precision != null)
                    rs.Fields[col.ColumnName].Precision = def.Precision.Value;
                if (def?.NumericScale != null)
                    rs.Fields[col.ColumnName].NumericScale = def.NumericScale.Value;
            }

            rs.Open();

            // Now, add the rows
            foreach (DataRow dr in table.Rows)
            {
                rs.AddNew();

                for (int x = 0; x < table.Columns.Count; x++)
                {
                    DataColumn col = table.Columns[x];
                    if (dr[col.Caption].GetType() == typeof(Guid))
                    {
                        rs.Fields[col.Caption].Value = ((Guid)dr[col.Caption]).ToString("B"); //need to format the GUID to a format that ABODB will accept
                    }
                    else
                    {
                        rs.Fields[col.Caption].Value = dr[col.Caption];
                    }
                }
                rs.Update();
            }

            return rs;
        }

        /// <summary>
        /// Builds a Recordset from a IDataReader object (e.g. SqlDataReader).
        /// </summary>
        /// <param name="reader">The source IDataReader.</param>
        /// <returns>A Recordset</returns>
        [Obsolete("Not acutally implemented yet.")]
        public static Recordset FromDataReader(IDataReader reader)
        {
            throw new NotImplementedException("Implement me");
        }

        /// <summary>
        /// Gets a collection of Field objects. This object is internally cached and does not need to be manually disposed.
        /// </summary>
        public Fields Fields
        {
            get
            {
                if (_fields == null)
                {
                    _fields = new Fields(InternalObject.GetType().InvokeMember("Fields", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _fields;
            }
        }

        /// <summary>
        /// Opens the RecordSet.
        /// This should be done after the record set is prepared, but before data is inserted into it/queried from it.
        /// </summary>
        public void Open()
        {
            InternalObject.GetType().InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        // TODO: Figure out why this doesn't work
        /*public void AddNew(string[] fields, object[] values)
        {
            _recordset.GetType().InvokeMember("AddNew", System.Reflection.BindingFlags.InvokeMethod, null, _recordset, new object[] { fields, values });
        }*/

        /// <summary>
        /// Shortcut for adding new a new row to the recordset.
        /// </summary>
        /// <typeparam name="V">Any valid data type.</typeparam>
        /// <param name="keyValuePairs">A dictionary of key-value pairs, where keys are the field names into which to insert data. Fields must already exist.</param>
        public void AddNew<V>(Dictionary<string, V> keyValuePairs)
        {
            AddNew();
            foreach (var pair in keyValuePairs)
            {
                Fields[pair.Key].Value = pair.Value;
            }
            Update();
        }

        /// <summary>
        /// Shortcut for adding new a new row to the recordset.
        /// </summary>
        /// <typeparam name="V">Any valid data type.</typeparam>
        /// <param name="fields">A list of field names into which to insert data. Fields must already exist.</param>
        /// <param name="values">A list of values corresponding in ordinal sequence to the list of field names.</param>
        public void AddNew<V>(IEnumerable<string> fields, IEnumerable<V> values)
        {
            AddNew();
            IEnumerator<string> field = fields.GetEnumerator();
            IEnumerator<V> value = values.GetEnumerator();
            while (field.MoveNext() && value.MoveNext())
            {
                Fields[field.Current].Value = value.Current;
            }
            Update();
        }

        // TODO: Get proper summary from MSDN.
        /// <summary>
        /// Begin adding a new row.
        /// </summary>
        public void AddNew()
        {
            InternalObject.GetType().InvokeMember("AddNew", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        // TODO: Get proper summary from MSDN.
        /// <summary>
        /// End adding a new row.
        /// </summary>
        public void Update()
        {
            InternalObject.GetType().InvokeMember("Update", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_fields != null)
                    _fields.Dispose();
            }

            base.Dispose(true);
        }

        #endregion
    }
}
