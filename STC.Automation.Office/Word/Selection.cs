using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using STC.Automation.Office.Word.Enums;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Represents the current selection in a window or pane.
    /// A selection represents either a selected (or highlighted) area in the document, or it represents the insertion point if nothing in the document is selected.
    /// There can be only one Selection object per document window pane, and only one Selection object in the entire application can be active.
    /// </summary>
    [WrapsCOM("Word.Selection", "00020975-0000-0000-C000-000000000046")]
    public class Selection : ComWrapper
    {
        internal Selection(object selectionObj)
            : base(selectionObj)
        {
        }

        /// <summary>
        /// Returns a Cells collection that represents the table cells in a selection.
        /// This object must be manually disposed.
        /// </summary>
        public Cells Cells
        {
            get
            {
                return new Cells(InternalObject.GetType().InvokeMember("Cells", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Gets a Columns collection that represents all the table columns in a selection.
        /// This object must be manually disposed.
        /// </summary>
        public Columns Columns
        {
            get
            {
                return new Columns(InternalObject.GetType().InvokeMember("Columns", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Gets a Find object that contains the criteria for a find operation.
        /// This object must be manually disposed.
        /// </summary>
        public Find Find
        {
            get
            {
                return new Find(InternalObject.GetType().InvokeMember("Find", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Gets or sets a Font object that represents the character formatting of the specified object.
        /// This object must be manually disposed.
        /// </summary>
        public Font Font
        {
            get
            {
                return new Font(InternalObject.GetType().InvokeMember("Font", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("Font", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Returns a FormFields collection that represents all the form fields in the selection.
        /// This object must be manually disposed.
        /// </summary>
        public FormFields FormFields
        {
            get
            {
                return new FormFields(InternalObject.GetType().InvokeMember("FormFields", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Inserts a page, column, or section break.
        /// When you insert a page or column break, the break replaces the selection. If you don't want to replace the selection, use the Collapse method before using the InsertBreak method.
        /// </summary>
        /// <param name="type"></param>
        public void InsertBreak(BreakType type)
        {
            InternalObject.GetType().InvokeMember("InsertBreak", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { type });
        }

        /// <summary>
        /// Inserts rows above the current selection.
        /// Microsoft Word inserts as many rows as there are in the current selection.
        /// To use this method, the current selection must be in a table.
        /// </summary>
        public void InsertRowsAbove()
        {
            InternalObject.GetType().InvokeMember("InsertRowsAbove", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Inserts rows below the current selection.
        /// Microsoft Word inserts as many rows as there are in the current selection.
        /// To use this method, the current selection must be in a table.
        /// </summary>
        public void InsertRowsBelow()
        {
            InternalObject.GetType().InvokeMember("InsertRowsBelow", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Moves the selection down and returns the number of units it has been moved. Defaults to moving the selection 1 line.
        /// </summary>
        /// <returns>Number of units the selection has been moved.</returns>
        public long MoveDown()
        {
            return MoveDown(Units.Line, 1, MovementType.Move);
        }

        /// <summary>
        /// Moves the selection down and returns the number of units it has been moved.
        /// </summary>
        /// <param name="unit">The unit by which the selection is to be moved.</param>
        /// <param name="count">The number of units the selection is to be moved.</param>
        /// <param name="movementType">Can be either Move or Extend. If Move is used, the selection is collapsed to the endpoint and moved down.
        ///     If Extend is used, the selection is extended down.</param>
        /// <returns>Number of units the selection has been moved.</returns>
        public long MoveDown(Units unit, int count, MovementType movementType)
        {
            return Convert.ToInt64(InternalObject.GetType().InvokeMember("MoveDown", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { unit, count, movementType }));
        }

        /// <summary>
        /// Moves the selection to the right and returns the number of units it has been moved. Defaults to moving the selection 1 character.
        /// </summary>
        /// <returns>Number of units the selection has been moved.</returns>
        public long MoveRight()
        {
            return MoveRight(Units.Character, 1, MovementType.Move);
        }

        /// <summary>
        /// Moves the selection to the right and returns the number of units it has been moved.
        /// When the Unit is Cell, the Extend argument can only be Move.
        /// </summary>
        /// <param name="unit">The unit by which the selection is to be moved.</param>
        /// <param name="count">The number of units the selection is to be moved.</param>
        /// <param name="movementType">Can be either Move or Extend. If Move is used, the selection is collapsed to the endpoint and moved to the right.
        ///     If Extend is used, the selection is extended down.</param>
        /// <remarks>When the Unit is Cell, the Extend argument can only be Move.</remarks>
        /// <returns>Number of units the selection has been moved.</returns>
        public long MoveRight(Units unit, int count, MovementType movementType)
        {
            return Convert.ToInt64(InternalObject.GetType().InvokeMember("MoveRight", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { unit, count, movementType }));
        }

        /// <summary>
        /// Gets or sets a ParagraphFormat object that represents the paragraph settings for the specified selection.
        /// This object must be manually disposed.
        /// </summary>
        public ParagraphFormat ParagraphFormat
        {
            get
            {
                return new ParagraphFormat(InternalObject.GetType().InvokeMember("ParagraphFormat", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }

            set
            {
                InternalObject.GetType().InvokeMember("ParagraphFormat", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value.InternalObject });
            }
        }

        /// <summary>
        /// Gets a Range object that represents the portion of a document that's contained in the selection.
        /// This object must be manually disposed.
        /// </summary>
        public Range Range
        {
            get
            {
                return new Range(InternalObject.GetType().InvokeMember("Range", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Gets a Rows collection that represents all the table rows in a selection.
        /// This object must be manually disposed.
        /// </summary>
        public Rows Rows
        {
            get
            {
                return new Rows(InternalObject.GetType().InvokeMember("Rows", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Inserts a new, blank paragraph.
        /// This method corresponds to the functionality of the ENTER key. If the selection isn't collapsed to an insertion point, the new paragraph replaces the selection.
        /// Use the InsertParagraphAfter or InsertParagraphBefore method to insert a new paragraph without deleting the contents of the selection.
        /// </summary>
        public void TypeParagraph()
        {
            InternalObject.GetType().InvokeMember("TypeParagraph", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Inserts the specified text.
        /// If the Word.Options.ReplaceSelection property is True, the selection is replaced by the specified text. If ReplaceSelection is False, the specified text is inserted before the selection.
        /// </summary>
        /// <param name="text">The text to be inserted.</param>
        public void TypeText(string text)
        {
            InternalObject.GetType().InvokeMember("TypeText", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { text });
        }
    }
}
