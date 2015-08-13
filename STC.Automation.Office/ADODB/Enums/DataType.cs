using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.ADODB.Enums
{
    /// <summary>
    /// Specifies the data type of a Field, Parameter, or Property.
    /// </summary>
    public enum DataType
    {
        /// <summary>
        /// Specifies no value (DBTYPE_EMPTY).
        /// </summary>
        Empty = 0,
        /// <summary>
        /// Indicates a two-byte signed integer (DBTYPE_I2).
        /// </summary>
        SmallInt = 2,
        /// <summary>
        /// Indicates a four-byte signed integer (DBTYPE_I4).
        /// </summary>
        Integer = 3,
        /// <summary>
        /// Indicates a single-precision floating-point value (DBTYPE_R4).
        /// </summary>
        Single = 4,
        /// <summary>
        /// Indicates a double-precision floating-point value (DBTYPE_R8).
        /// </summary>
        Double = 5,
        /// <summary>
        /// Indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an eight-byte signed integer scaled by 10,000.
        /// </summary>
        Currency = 6,
        /// <summary>
        /// Indicates a date value (DBTYPE_DATE). A date is stored as a double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day.
        /// </summary>
        Date = 7,
        /// <summary>
        /// Indicates a null-terminated character string (Unicode) (DBTYPE_BSTR).
        /// </summary>
        BSTR = 8,
        /// <summary>
        /// Indicates a pointer to an IDispatch interface on a COM object (DBTYPE_IDISPATCH). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results.
        /// </summary>
        IDispatch = 9,
        /// <summary>
        /// Indicates a 32-bit error code (DBTYPE_ERROR).
        /// </summary>
        Error = 10,
        /// <summary>
        /// Indicates a Boolean value (DBTYPE_BOOL).
        /// </summary>
        Boolean = 11,
        /// <summary>
        /// Indicates an Automation Variant (DBTYPE_VARIANT). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results.
        /// </summary>
        Variant = 12,
        /// <summary>
        /// Indicates a pointer to an IUnknown interface on a COM object (DBTYPE_IUNKNOWN). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results.
        /// </summary>
        IUnknown = 13,
        /// <summary>
        /// Indicates an exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL).
        /// </summary>
        Decimal = 14,
        /// <summary>
        /// Indicates a one-byte signed integer (DBTYPE_I1).
        /// </summary>
        TinyInt = 16,
        /// <summary>
        /// Indicates a one-byte unsigned integer (DBTYPE_UI1).
        /// </summary>
        UnsignedTinyInt = 17,
        /// <summary>
        /// Indicates a two-byte unsigned integer (DBTYPE_UI2).
        /// </summary>
        UnsignedSmallInt = 18,
        /// <summary>
        /// Indicates a four-byte unsigned integer (DBTYPE_UI4).
        /// </summary>
        UnsignedInt = 19,
        /// <summary>
        /// Indicates an eight-byte signed integer (DBTYPE_I8).
        /// </summary>
        BigInt = 20,
        /// <summary>
        /// Indicates an eight-byte unsigned integer (DBTYPE_UI8).
        /// </summary>
        UnsignedBigInt = 21,
        /// <summary>
        /// Indicates a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (DBTYPE_FILETIME).
        /// </summary>
        FileTime = 64,
        /// <summary>
        /// Indicates a globally unique identifier (GUID) (DBTYPE_GUID).
        /// </summary>
        GUID = 72,
        /// <summary>
        /// Indicates a binary value (DBTYPE_BYTES).
        /// </summary>
        Binary = 128,
        /// <summary>
        /// Indicates a string value (DBTYPE_STR).
        /// </summary>
        Char = 129,
        /// <summary>
        /// Indicates a null-terminated Unicode character string (DBTYPE_WSTR).
        /// </summary>
        WChar = 130,
        /// <summary>
        /// Indicates an exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC).
        /// </summary>
        Numeric = 131,
        /// <summary>
        /// Indicates a user-defined variable (DBTYPE_UDT).
        /// </summary>
        UserDefined = 132,
        /// <summary>
        /// Indicates a date value (yyyymmdd) (DBTYPE_DBDATE).
        /// </summary>
        DBDate = 133,
        /// <summary>
        /// Indicates a time value (hhmmss) (DBTYPE_DBTIME).
        /// </summary>
        DBTime = 134,
        /// <summary>
        /// Indicates a date/time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP).
        /// </summary>
        DBTimeStamp = 135,
        /// <summary>
        /// Indicates a four-byte chapter value that identifies rows in a child rowset (DBTYPE_HCHAPTER).
        /// </summary>
        Chapter = 136,
        /// <summary>
        /// Indicates an Automation PROPVARIANT (DBTYPE_PROP_VARIANT).
        /// </summary>
        PropVariant = 138,
        /// <summary>
        /// Indicates a numeric value.
        /// </summary>
        VarNumeric = 139,
        /// <summary>
        /// Indicates a string value.
        /// </summary>
        VarChar = 200,
        /// <summary>
        /// Indicates a long string value.
        /// </summary>
        LongVarChar = 201,
        /// <summary>
        /// Indicates a null-terminated Unicode character string.
        /// </summary>
        VarWChar = 202,
        /// <summary>
        /// Indicates a long null-terminated Unicode string value.
        /// </summary>
        LongVarWChar = 203,
        /// <summary>
        /// Indicates a binary value.
        /// </summary>
        VarBinary = 204,
        /// <summary>
        /// Indicates a long binary value.
        /// </summary>
        LongVarBinary = 205,
        /// <summary>
        /// A flag value, always combined with another data type constant, that indicates an array of the other data type. Does not apply to ADOX.
        /// </summary>
        Array = 8192,
    }
}
