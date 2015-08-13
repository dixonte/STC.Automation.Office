using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Excel.Enums;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Describes a sub-section of cells within a worksheet
    /// </summary>
    [WrapsCOM("Excel.Range", "00020846-0000-0000-C000-000000000046")]
    public class Range : ComWrapper
    {
        private Application _application;
        private Worksheet _worksheet;
        private Range _columns;
        private Range _rows;
        private Range _entireColumn;
        private Range _entireRow;
        private Borders _borders;
        private Font _font;
        private Comment _comment;

        internal Range(object rangeObj)
            : base(rangeObj)
        {
        }

        /// <summary>
        /// Activates a single cell, which must be inside the current selection. To select a range of cells, use the Select method.
        /// </summary>
        public void Activate()
        {
            InternalObject.GetType().InvokeMember("Activate", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Filters a list using the AutoFilter.
        /// </summary>
        /// <param name="field">The integer offset of the field on which you want to base the filter (from the left of the list; the leftmost field is field one).</param>
        /// <param name="criteria1">The criteria (a string; for example, "101"). Use "=" to find blank fields, or use "<>" to find nonblank fields. If this argument is omitted, the criteria is All. If Operator is xlTop10Items, Criteria1 specifies the number of items (for example, "10").</param>
        /// <param name="criteriaOperator">One of the constants of XlAutoFilterOperator specifying the type of filter.</param>
        /// <param name="criteria2">The second criteria (a string). Used with Criteria1 and Operator to construct compound criteria.</param>
        /// <param name="visibleDropDown">True to display the AutoFilter drop-down arrow for the filtered field. False to hide the AutoFilter drop-down arrow for the filtered field. True by default.</param>
        public void AutoFilter(int? field = null, string criteria1 = null, AutoFilterOperator? criteriaOperator = null, string criteria2 = null, bool? visibleDropDown = null)
        {
            InternalObject.GetType().InvokeMember("AutoFilter", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, ComArguments.Prepare(field, criteria1, criteriaOperator, criteria2, visibleDropDown));
        }

        /// <summary>
        /// Sorts a range of values
        /// </summary>
        /// <param name="key1">Specifies the first sort field, either as a range name (string) or Range object; determines the values to be sorted.</param>
        /// <param name="order1">Determines the sort order for the values specified in key1.</param>
        /// <param name="type">Specifies which elements are to be sorted.</param>
        /// <param name="key2">Second sort field;l cannot be used when sorting a pivot table.</param>
        /// <param name="order2">Determines the sort order for the values specified in Key2.</param>
        /// <param name="key3">Third sort field; cannot be used when sorting a pivot table.</param>
        /// <param name="order3">Determines the sort order for the values specified in Key3.</param>
        /// <param name="header">Specifies whether the first row contains header information. xlNo is the default value; specify xlGuess if you want Excel to attempt to determine the header.</param>
        /// <returns></returns>
        [System.Obsolete("This method has not been fully tested yet and is not guaranteed to work")]
        public void Sort(Range key1=null, SortOrder? order1=null, Range key2 = null, SortType? type = null, SortOrder? order2 = null, Range key3 = null, SortOrder? order3 = null, YesNoGuess? header = YesNoGuess.No)
        {
            //if (!(key1 is String) && !(key1 is Range) && key1 != null)
            //    throw new ArgumentException("Key1 must be a string (range named) or a range object");

            //if (!(key2 is String) && !(key2 is Range) && key2 != null)
            //    throw new ArgumentException("Key2 must be a string (range named) or a range object");

            //if (!(key3 is String) && !(key3 is Range) && key3 != null)
            //    throw new ArgumentException("Key3 must be a string (range named) or a range object");

            InternalObject.GetType().InvokeMember("Sort", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, ComArguments.Prepare(key1, order1, key2, order2, key3, order3, header));
        }


        /// <summary>
        /// Returns a String value that represents the range reference in the language of the macro
        /// </summary>
        /// <remarks>If the reference contains more than one cell, RowAbsolute and ColumnAbsolute apply to all rows and columns.</remarks>
        /// <param name="rowAbsolute">True to return the row part of the reference as an absolute reference. The default value is True.</param>
        /// <param name="columnAbsolute">True to return the column part of the reference as an absolute reference. The default value is True.</param>
        /// <param name="referenceStyle">Specifies the reference style.</param>
        /// <param name="external">True to return an external reference; False to return a local reference. The default value is False.</param>
        /// <param name="relativeTo">If RowAbsolute and ColumnAbsolute are False, and ReferenceStyle is xlR1C1, you must include a starting point for the relative reference. This argument is a Range object that defines the starting point.</param>
        /// <returns>The address</returns>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public string Address(bool rowAbsolute = true, bool columnAbsolute = true, XLReferenceStyle referenceStyle = XLReferenceStyle.xlA1, bool external = false, Range relativeTo = null)
        {
            return InternalObject.GetType().InvokeMember("Address", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { rowAbsolute, columnAbsolute, referenceStyle, external, (relativeTo == null) ? System.Reflection.Missing.Value : relativeTo.InternalObject }).ToString();
        }

        /// <summary>
        /// Find
        /// </summary>
        /// <param name="what">String - what to find</param>
        /// <returns>a range object</returns>
        [System.Obsolete("This method has not been fully tested yet and is not guaranteed to work")]
        public Range Find(string what)
        {
            return (Range)InternalObject.GetType().InvokeMember("Find", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { what });
        }

        /// <summary>
        /// Copies the range to the specified range or to the Clipboard.
        /// </summary>
        /// <param name="Destination">Specifies the new range to which the specified range will be copied. If this argument is omitted, Microsoft Excel copies the range to the Clipboard.</param>
        [System.Obsolete("This method has not been fully tested yet and is not guaranteed to work")]
        public void Copy(Range Destination = null)
        {
            InternalObject.GetType().InvokeMember("Copy", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { Destination });
        }

        /// <summary>
        /// Gets a reference to the currently worksheet to which the range belongs. The returned worksheet is internally cached and does not need to be manually disposed.
        /// </summary>
        public Application Application
        {
            get
            {
                if (_application == null)
                    _application = new Application(InternalObject.GetType().InvokeMember("Application", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _application;
            }
        }

        /// <summary>
        /// Adds a comment to the range.
        /// </summary>
        public void AddComment(string text)
        {
            InternalObject.GetType().InvokeMember("AddComment", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { text });
        }

        /// <summary>
        /// Clears all cell comments from the specified range.
        /// </summary>
        public void ClearComments()
        {
            InternalObject.GetType().InvokeMember("ClearComments", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Gets a Comment object that represents the comment associated with the cell in the upper-left corner of the range. This object is internally cached and does not require manual disposal.
        /// </summary>
        public Comment Comment
        {
            get
            {
                if (_comment == null)
                {
                    object obj = InternalObject.GetType().InvokeMember("Comment", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
                    if (obj != null)
                        _comment = new Comment(obj);
                }

                return _comment;
            }
        }

        

        /// <summary>
        /// Changes the width of the columns in the range or the height of the rows in the range to achieve the best fit.
        /// </summary>
        public void AutoFit()
        {
            InternalObject.GetType().InvokeMember("AutoFit", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Gets a Borders collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format). This object is internally cached and does not require manual disposal.
        /// </summary>
        public Borders Borders
        {
            get
            {
                if (_borders == null)
                    _borders = new Borders(InternalObject.GetType().InvokeMember("Borders", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _borders;
            }
        }

        /// <summary>
        /// Gets a Range object that represents the columns in the specified range. This Range is internally cached and does not require manual disposal.
        /// </summary>
        public Range Columns
        {
            get
            {
                if (_columns == null)
                    _columns = new Range(InternalObject.GetType().InvokeMember("Columns", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _columns;
            }
        }

        /// <summary>
        /// Gets a Range object that represents the rows in the specified range. This Range is internally cached and does not require manual disposal.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public Range Rows
        {
            get
            {
                if (_rows == null)
                    _rows = new Range(InternalObject.GetType().InvokeMember("Rows", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _rows;
            }
        }

        /// <summary>
        /// Returns a value that represents the number of objects in the collection (e.g. row count or column count depending on the range).
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public int Count
        {
            get
            {
                object result = InternalObject.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);

                if (result != null && result != DBNull.Value)
                {
                    return Convert.ToInt32(result);
                }

                return 1;
            }
        }

        /// <summary>
        /// Returns the number of the first column in the first area in the specified range.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public int Column
        {
            get
            {
                object result = InternalObject.GetType().InvokeMember("Column", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);

                if (result != null && result != DBNull.Value)
                {
                    return Convert.ToInt32(result);
                }

                return 1;
            }
        }

        /// <summary>
        /// Returns the number of the first row of the first area in the range.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public int Row
        {
            get
            {
                object result = InternalObject.GetType().InvokeMember("Row", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);

                if (result != null && result != DBNull.Value)
                {
                    return Convert.ToInt32(result);
                }

                return 1;
            }
        }

        /// <summary>
        /// Gets or sets the width of all columns in the specified range, in units of average character width in the Normal style.
        /// </summary>
        /// <remarks>
        /// One unit of column width is equal to the width of one character in the Normal style. For proportional fonts, the width of the character 0 (zero) is used.
        /// Use the Width property to return the width of a column in points.
        /// If all columns in the range have the same width, the ColumnWidth property returns the width. If columns in the range have different widths, this property returns null.
        /// </remarks>
        public Decimal? ColumnWidth
        {
            get
            {
                object result = InternalObject.GetType().InvokeMember("ColumnWidth", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);

                if (result != null && result != DBNull.Value)
                {
                    return Convert.ToDecimal(result);
                }

                return null;
            }

            set
            {
                InternalObject.GetType().InvokeMember("ColumnWidth", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }


        /// <summary>
        /// Gets or sets the height of all the rows in the range specified, measured in points. Returns null if the rows in the specified range aren't all the same height.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public Decimal? RowHeight
        {
            get
            {
                object result = InternalObject.GetType().InvokeMember("RowHeight", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);

                if (result != null && result != DBNull.Value)
                {
                    return Convert.ToDecimal(result);
                }

                return null;
            }

            set
            {
                InternalObject.GetType().InvokeMember("RowHeight", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }

        }
        

        /// <summary>
        /// Copies the contents of an ADO Recordset object onto a worksheet, beginning at the upper-left corner of the specified range.
        /// If the Recordset object contains fields with OLE objects in them, this method fails.
        /// </summary>
        /// <param name="recordset">A wrapped ADODB.Recordset object representing the data to copy</param>
        /// <returns>?</returns>
        public int CopyFromRecordset(ADODB.Recordset recordset)
        {
            return CopyFromRecordset(recordset.InternalObject);
        }

        /// <summary>
        /// Copies the contents of an ADO Recordset object onto a worksheet, beginning at the upper-left corner of the specified range.
        /// If the Recordset object contains fields with OLE objects in them, this method fails.
        /// </summary>
        /// <param name="recordset">A wrapped ADODB.Recordset object representing the data to copy</param>
        /// <param name="maxRows">The maximum number of records to copy onto the worksheet. If this argument is omitted, all the records in the Recordset object are copied.</param>
        /// <param name="maxColumns">The maximum number of fields to copy onto the worksheet. If this argument is omitted, all the fields in the Recordset object are copied.</param>
        /// <returns>?</returns>
        public int CopyFromRecordset(ADODB.Recordset recordset, int? maxRows, int? maxColumns)
        {
            return CopyFromRecordset(recordset.InternalObject, maxRows, maxColumns);
        }

        /// <summary>
        /// Copies the contents of an ADO Recordset object onto a worksheet, beginning at the upper-left corner of the specified range.
        /// If the Recordset object contains fields with OLE objects in them, this method fails.
        /// </summary>
        /// <param name="recordset">A raw ADODB.Recordset or DAO.Recordset COM object representing the data to copy</param>
        /// <returns>?</returns>
        public int CopyFromRecordset(object recordset)
        {
            return CopyFromRecordset(recordset, null, null);
        }

        /// <summary>
        /// Copies the contents of an ADO Recordset object onto a worksheet, beginning at the upper-left corner of the specified range.
        /// If the Recordset object contains fields with OLE objects in them, this method fails.
        /// </summary>
        /// <param name="recordset">A raw ADODB.Recordset or DAO.Recordset COM object representing the data to copy</param>
        /// <param name="maxRows">The maximum number of records to copy onto the worksheet. If this argument is omitted, all the records in the Recordset object are copied.</param>
        /// <param name="maxColumns">The maximum number of fields to copy onto the worksheet. If this argument is omitted, all the fields in the Recordset object are copied.</param>
        /// <returns>?</returns>
        public int CopyFromRecordset(object recordset, int? maxRows, int? maxColumns)
        {
            List<object> parms = new List<object>();
            parms.Add(recordset);
            if (maxRows != null)
                parms.Add(maxRows);
            if (maxColumns != null)
            {
                while (parms.Count < 2)
                {
                    parms.Add(System.Reflection.Missing.Value);
                }
                parms.Add(maxColumns);
            }

            return Convert.ToInt32(InternalObject.GetType().InvokeMember("CopyFromRecordset", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, parms.ToArray()));
        }

        /// <summary>
        /// Deletes the object.
        /// </summary>
        public void Delete()
        {
            Delete(null);
        }

        /// <summary>
        /// Deletes the object.
        /// </summary>
        /// <param name="shift">Used only with Range objects. Specifies how to shift cells to replace deleted cells. Can be one of the following DeleteShiftDirection constants: ToLeft or Up.
        /// If this argument is omitted, Microsoft Excel decides based on the shape of the range.</param>
        public void Delete(DeleteShiftDirection? shift)
        {
            List<object> parms = new List<object>();
            if (shift != null)
                parms.Add(shift.Value);

            InternalObject.GetType().InvokeMember("Delete", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, parms.ToArray());
        }

        /// <summary>
        /// Clears the entire object.
        /// </summary>
        public void Clear()
        {
            InternalObject.GetType().InvokeMember("Clear", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Adds a border to a range and sets the Color, LineStyle, and Weight properties for the new border
        /// </summary>
        /// <param name="linestyle">One of the constants of LineStyle specifying the line style for the border.</param>
        /// <param name="borderweight">The border weight</param>
        /// <param name="colourindex">The border color, as an index into the current color palette or as a ColorIndex constant.</param>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public void BorderAround(LineStyle? linestyle = null, BorderWeight? borderweight = null, ColorIndex? colourindex = null)
        {
            InternalObject.GetType().InvokeMember("BorderAround", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] {linestyle, borderweight, colourindex });
        }

        /// <summary>
        /// Gets a Range object that represents the entire column (or columns) that contains the specified range. This Range is internally cached and does not require manual disposal.
        /// </summary>
        public Range EntireColumn
        {
            get
            {
                if (_entireColumn == null)
                    _entireColumn = new Range(InternalObject.GetType().InvokeMember("EntireColumn", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _entireColumn;
            }
        }

        /// <summary>
        /// Returns a Range object that represents the entire row (or rows) that contains the specified range. This Range is internally cached and does not require manual disposal.
        /// </summary>
        public Range EntireRow
        {
            get
            {
                if (_entireRow == null)
                    _entireRow = new Range(InternalObject.GetType().InvokeMember("EntireRow", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _entireRow;
            }
        }

        /// <summary>
        /// Gets a Font object that represents the font of the specified Range. This Font is internally cached and does not require manual disposal.
        /// </summary>
        public Font Font
        {
            get
            {
                if (_font == null)
                    _font = new Font(InternalObject.GetType().InvokeMember("Font", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _font;
            }
        }

        /// <summary>
        /// Gets or sets the formula for the object, using R1C1-style notation in the language of the macro.
        /// </summary>
        /// <remarks>
        /// If the cell contains a constant, this property returns the constant. If the cell is empty, the property returns an empty string. If the cell contains a formula, the property returns the formula as a string, in the same format in which it would be displayed in the formula bar (including the equal sign).
        /// If you set the value or formula of a cell to a date, Microsoft Excel checks to see whether that cell is already formatted with one of the date or time number formats. If not, the number format is changed to the default short date number format.
        /// If the range is a one- or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.
        /// Setting the formula of a multiple-cell range fills all cells in the range with the formula.
        /// </remarks>
        public string FormulaR1C1
        {
            get
            {
                return InternalObject.GetType().InvokeMember("FormulaR1C1", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null) as string;
            }

            set
            {
                InternalObject.GetType().InvokeMember("FormulaR1C1", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Get Range given a row and column index.  The returned Range must be manually disposed.
        /// </summary>
        /// <param name="row">Row index</param>
        /// <param name="column">Column index</param>
        /// <returns>The range at this row and column location.</returns>
        public Range GetRange(int column, int row)
        {
            return new Range(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, InternalObject, new object[] { row, column }));
        }

        /// <summary>
        /// Returns or sets a Variant value that represents the height, in points, of the range.
        /// </summary>
        public double Height
        {
            get
            {
                return Convert.ToDouble(InternalObject.GetType().InvokeMember("Height", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Gets or sets a Variant value that represents the horizontal alignment for the specified object.
        /// </summary>
        public Enums.HAlign HorizontalAlignment
        {
            get
            {
                return (Enums.HAlign)InternalObject.GetType().InvokeMember("HorizontalAlignment", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("HorizontalAlignment", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Get or sets a Varient value that represents the vertical alignment for the specified object.
        /// </summary>
        public Enums.VAlign VerticalAlignment
        {
            get
            {
                return (Enums.VAlign)InternalObject.GetType().InvokeMember("VerticalAlignment", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("VerticalAlignment", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns a Variant value that represents the distance, in points, from the left edge of column A to the left edge of the range.
        /// </summary>
        public double Left
        {
            get
            {
                return Convert.ToDouble(InternalObject.GetType().InvokeMember("Left", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Creates a merged cell from the specified Range object.
        /// </summary>
        public void Merge()
        {
            Merge(null);
        }

        /// <summary>
        /// Creates a merged cell from the specified Range object.
        /// </summary>
        /// <param name="across">True to merge cells in each row of the specified range as separate merged cells. The default value is False.</param>
        public void Merge(bool? across)
        {
            List<object> parms = new List<object>();
            if (across != null)
                parms.Add(across.Value);

            InternalObject.GetType().InvokeMember("Merge", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, parms.ToArray());
        }

        /// <summary>
        /// True if the range contains merged cells. Read/write boolean.
        /// </summary>
        /// <remarks>
        /// When you select a range that contains merged cells, the resulting selection may be different from the intended selection. Use the Address  property to check the address of the selected range.
        /// </remarks>
        public bool MergeCells
        {
            get
            {
                return Convert.ToBoolean(InternalObject.GetType().InvokeMember("MergeCells", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("MergeCells", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets a string value that represents the format code for the object.
        /// </summary>
        /// <remarks>
        /// This property returns Null if all cells in the specified range don't have the same number format.
        /// The format code is the same string as the Format Codes option in the Format Cells dialog box.
        /// The Format function uses different format code strings than do the NumberFormat and NumberFormatLocal properties.
        /// </remarks>
        public string NumberFormat
        {
            get
            {
                return InternalObject.GetType().InvokeMember("NumberFormat", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null) as string;
            }

            set
            {
                InternalObject.GetType().InvokeMember("NumberFormat", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Selects the range.
        /// </summary>
        public void Select()
        {
            InternalObject.GetType().InvokeMember("Select", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        

        /// <summary>
        /// Shortcut for GetRange(row, column).
        /// </summary>
        /// <param name="column">Column index</param>
        /// <param name="row">Row index</param>
        /// <returns>The range at the given row and column.</returns>
        public Range this[int column, int row]
        {
            get
            {
                Range r = GetRange(column, row);
                return r;
            }

            
        }

        

        /// <summary>
        /// Returns a Variant value that represents the distance, in points, from the top edge of row 1 to the top edge of the range.
        /// </summary>
        public double Top
        {
            get
            {
                return Convert.ToDouble(InternalObject.GetType().InvokeMember("Top", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Returns or sets an object value that represents the value of the specified range.  This probably doesn't need to be disposed.
        /// </summary>
        /// <remarks>When setting a range of cells with the contents of an XML spreadsheet file, only values of the first sheet in the workbook are used. You cannot set or get a discontiguous range of cells in the XML spreadsheet format.</remarks>
        public object Value
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Value", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }
            set
            {
                InternalObject.GetType().InvokeMember("Value", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the location of a page break. Can be one of the PageBreak constants. 
        /// This property can return the location of either automatic or manual page breaks, but it can only set the location of manual breaks (it can only be set to PageBreakManual or PageBreakNone).
        /// To remove all manual page breaks on a worksheet, set Cells.PageBreak to xlPageBreakNone.
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public PageBreak PageBreak
        {
            get
            {
                return (PageBreak)GetType().InvokeMember("PageBreak", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                if (value == Enums.PageBreak.PageBreakAutomatic)
                    throw new ArgumentException("You can only set the pagebreak property to Manual or None.");

                InternalObject.GetType().InvokeMember("PageBreak", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns a Variant value that represents the width, in units, of the range.
        /// </summary>
        public double Width
        {
            get
            {
                return Convert.ToDouble(InternalObject.GetType().InvokeMember("Width", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Gets a reference to the currently worksheet to which the range belongs. The returned worksheet is internally cached and does not need to be manually disposed.
        /// </summary>
        public Worksheet Worksheet
        {
            get
            {
                if (_worksheet == null)
                    _worksheet = new Worksheet(InternalObject.GetType().InvokeMember("Worksheet", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _worksheet;
            }
        }

        /// <summary>
        /// Gets or sets a boolean value that indicates if Microsoft Excel wraps the text in the object.
        /// </summary>
        public bool WrapText
        {
            get
            {
                return Convert.ToBoolean(InternalObject.GetType().InvokeMember("WrapText", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("WrapText", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the Interior of the range. (this is not automatically disposed)
        /// </summary>
        [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
        public Interior Interior
        {
            get
            {
                return new Interior(InternalObject.GetType().InvokeMember("Interior", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Interior", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] {value});
            }
        }



        internal override void Dispose(bool disposing)
        {
            if (_comment != null)
            {
                _comment.Dispose();
                _comment = null;
            }

            if (_font != null)
            {
                _font.Dispose();
                _font = null;
            }

            if (_borders != null)
            {
                _borders.Dispose();
                _borders = null;
            }

            if (_columns != null)
            {
                _columns.Dispose();
                _columns = null;
            }

            if (_rows != null)
            {
                _rows.Dispose();
                _rows = null;
            }

            if (_entireColumn != null)
            {
                _entireColumn.Dispose();
                _entireColumn = null;
            }

            if (_entireRow != null)
            {
                _entireRow.Dispose();
                _entireRow = null;
            }

            if (_worksheet != null)
            {
                _worksheet.Dispose();
                _worksheet = null;
            }

            if (_application != null)
            {
                _application.Dispose();
                _application = null;
            }

            base.Dispose(disposing);
        }
    }
}
