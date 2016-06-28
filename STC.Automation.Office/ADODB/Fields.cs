using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.ADODB.Enums;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.ADODB
{
    /// <summary>
    /// Wraps an ADODB.Fields object
    /// </summary>
    [WrapsCOM("ADODB.Fields", "00000564-0000-0010-8000-00AA006D2EA4")]
    public class Fields : ComWrapper
    {
        internal Fields(object fieldObj)
            : base(fieldObj)
        {
        }

        /// <summary>
        /// Adds a new Field to this collection.
        /// </summary>
        /// <param name="name">Name of the field to add</param>
        /// <param name="dataType">Data type of the field to add</param>
        public void Append(string name, DataType dataType)
        {
            Append(name, dataType, null, null);
        }

        /// <summary>
        /// Adds a new Field to this collection.
        /// </summary>
        /// <param name="name">Name of the field to add</param>
        /// <param name="dataType">Data type of the field to add</param>
        /// <param name="size">Size of the field to add</param>
        public void Append(string name, DataType dataType, long? size)
        {
            Append(name, dataType, size, null);
        }

        /// <summary>
        /// Adds a new Field to this collection.
        /// </summary>
        /// <param name="name">Name of the field to add</param>
        /// <param name="dataType">Data type of the field to add</param>
        /// <param name="fieldAttributes">Attributes on the field to add</param>
        public void Append(string name, DataType dataType, FieldAttribute fieldAttributes)
        {
            Append(name, dataType, null, fieldAttributes);
        }

        /// <summary>
        /// Adds a new Field to this collection.
        /// </summary>
        /// <param name="name">Name of the field to add</param>
        /// <param name="dataType">Data type of the field to add</param>
        /// <param name="size">Size of the field to add</param>
        /// <param name="fieldAttributes">Attributes on the field to add</param>
        public void Append(string name, DataType dataType, long? size, FieldAttribute? fieldAttributes)
        {
            List<object> parms = new List<object>();
            parms.Add(name);
            parms.Add(dataType);
            parms.Add(size);
            parms.Add(fieldAttributes);

            InternalObject.GetType().InvokeMember("Append", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, parms.ToArray());
        }

        /// <summary>
        /// Retrieves a field from the collection.
        /// This Field must be manually disposed.
        /// </summary>
        /// <param name="idx">Name of the field to retrieve.</param>
        /// <returns>Field</returns>
        public Field this[string idx]
        {
            get
            {
                return new Field(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { idx }));
            }
        }

        /// <summary>
        /// Retrieves a field from the collection.
        /// This Field must be manually disposed.
        /// </summary>
        /// <param name="idx">Index of the field to retrieve.</param>
        /// <returns>Field</returns>
        public Field this[int idx]
        {
            get
            {
                return new Field(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { idx }));
            }
        }
    }
}
