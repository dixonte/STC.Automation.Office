using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.ADODB.Enums
{
    /// <summary>
    /// Specifies one or more attributes of a Field object.
    /// </summary>
    public enum FieldAttribute
    {
        /// <summary>
        /// Indicates that the provider caches field values and that subsequent reads are done from the cache.
        /// </summary>
        CacheDeferred = 0x1000,
        /// <summary>
        /// Indicates that the field contains fixed-length data.
        /// </summary>
        Fixed = 0x10,
        /// <summary>
        /// Indicates that the field contains a chapter value, which specifies a specific child recordset related to this parent field. Typically chapter fields are used with data shaping or filters.
        /// </summary>
        IsChapter = 0x2000,
        /// <summary>
        /// Indicates that the field specifies that the resource represented by the record is a collection of other resources, such as a folder, rather than a simple resource, such as a text file.
        /// </summary>
        IsCollection = 0x40000,
        /// <summary>
        /// Indicates that the field contains the default stream for the resource represented by the record. For example, the default stream can be the HTML content of a root folder on a Web site, which is automatically served when the root URL is specified.
        /// </summary>
        IsDefaultStream = 0x20000,
        /// <summary>
        /// Indicates that the field accepts null values.
        /// </summary>
        IsNullable = 0x20,
        /// <summary>
        /// Indicates that the field contains the URL that names the resource from the data store represented by the record. 
        /// </summary>
        IsRowURL = 0x10000,
        /// <summary>
        /// Indicates that the field specifies all or part of the column's primary key.
        /// </summary>
        KeyColumn = 0x8000,
        /// <summary>
        /// Indicates that the field is a long binary field. Also indicates that you can use the AppendChunk and GetChunk methods.
        /// </summary>
        Long = 0x80,
        /// <summary>
        /// Indicates that you can read null values from the field.
        /// </summary>
        MayBeNull = 0x40,
        /// <summary>
        /// Indicates that the field is deferred—that is, the field values are not retrieved from the data source with the whole record, but only when you explicitly access them.
        /// </summary>
        MayDefer = 2,
        /// <summary>
        /// Indicates that the field represents a numeric value from a column that supports negative scale values. The scale is specified by the NumericScale property.
        /// </summary>
        NegativeScale = 0x4000,
        /// <summary>
        /// Indicates that the field contains a persistent row identifier that cannot be written to and has no meaningful value except to identify the row (such as a record number, unique identifier, and so forth).
        /// </summary>
        RowID = 0x100,
        /// <summary>
        /// Indicates that the field contains some kind of time or date stamp used to track updates.
        /// </summary>
        RowVersion = 0x200,
        /// <summary>
        /// Indicates that the provider cannot determine if you can write to the field.
        /// </summary>
        UnknownUpdatable = 8,
        /// <summary>
        /// Indicates that the provider does not specify the field attributes.
        /// </summary>
        Unspecified = -1,
        /// <summary>
        /// Indicates that you can write to the field.
        /// </summary>
        Updatable = 4
    }
}
