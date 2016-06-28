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
    /// Wraps an ADODB.Field object
    /// </summary>
    [WrapsCOM("ADODB.Field", "00000569-0000-0010-8000-00AA006D2EA4")]
    public class Field : ComWrapper
    {
        internal Field(object fieldObj)
            : base(fieldObj)
        {
        }

        /// <summary>
        /// Gets or sets a value that indicates the number of decimal places to which numeric values will be resolved.
        /// </summary>
        public byte NumericScale
        {
            get
            {
                return (byte)InternalObject.GetType().InvokeMember("NumericScale", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("NumericScale", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets a value that indicates the maximum number of digits used to represent values.
        /// </summary>
        public byte Precision
        {
            get
            {
                return (byte)InternalObject.GetType().InvokeMember("Precision", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("Precision", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets the value stored in this field.
        /// If the returned object is a COM object, you will have to clean it up yourself.
        /// </summary>
        public object Value
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Value", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                object val;

                if (value != null)
                    val = value;
                else
                    val = DBNull.Value;

                InternalObject.GetType().InvokeMember("Value", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { val });
            }
        }

        /// <summary>
        /// Converts this field to a string representation.
        /// </summary>
        /// <returns>String</returns>
        public override string ToString()
        {
            return Value != null ? Value.ToString() : "NULL";
        }
    }
}
