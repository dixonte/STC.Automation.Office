using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Word.Enums;

namespace STC.Automation.Office.Word
{
    /// <summary>
    /// Wraps an Word._Document object
    /// </summary>
    [WrapsCOM("Word._Document", "0002096B-0000-0000-C000-000000000046")]
    public class Document : ComWrapper
    {
        private Sections _sections;
        private Bookmarks _bookmarks;
        private ContentControls _contentControls;

        internal Document(object workbooksObj)
            : base(workbooksObj)
        {
        }

        /// <summary>
        /// Activates the specified document so that it becomes the active document.
        /// </summary>
        public void Activate()
        {
            InternalObject.GetType().InvokeMember("Activate", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Returns a Window object that represents the active window (the window with the focus). Read-only. You must dispose of this object manually.
        /// </summary>
        public Window ActiveWindow
        {
            get
            {
                return new Window(InternalObject.GetType().InvokeMember("ActiveWindow", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Gets a Bookmarks collection that represents all the bookmarks in a document. This item is internally cached and does not require manual disposal.
        /// </summary>
        public Bookmarks Bookmarks
        {
            get
            {
                if (_bookmarks == null)
                {
                    _bookmarks = new Bookmarks(InternalObject.GetType().InvokeMember("Bookmarks", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _bookmarks;
            }
        }

        /// <summary>
        /// Closes the document.
        /// </summary>
        public void Close()
        {
            Close(null);
        }

        /// <summary>
        /// Closes the document.
        /// </summary>
        /// <param name="saveChanges">Saves changes if true, abandons them if false, and asks the user if null</param>
        public void Close(bool? saveChanges)
        {
            List<object> parms = new List<object>();
            if (saveChanges != null)
                parms.Add(saveChanges.Value);

            InternalObject.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, parms.ToArray());

            Dispose(true);
        }

        /// <summary>
        /// Returns a statistic based on the contents of the specified document.
        /// </summary>
        /// <param name="statistic">The statistic to compute.</param>
        /// <param name="includeFootnotesAndEndnotes">True to include footnotes and endnotes when computing statistics.</param>
        /// <returns>long</returns>
        public long ComputeStatistics(Statistic statistic, bool includeFootnotesAndEndnotes)
        {
            return Convert.ToInt64(InternalObject.GetType().InvokeMember("ComputeStatistics", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { statistic, includeFootnotesAndEndnotes }));
        }

        /// <summary>
        /// Returns a Range object that represents the main document story. Read-only.
        /// This item is NOT internally cached and does require manual disposal.
        /// </summary>
        public Range Content
        {
            get
            {
                return new Range(InternalObject.GetType().InvokeMember("Content", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Gets a ContentControls collection that represents all the Content Controls in a document.
        /// This collection object is internally cached and does not need to be manually disposed.
        /// If enumerating this object using foreach(), you must manually dispose every instance you enumerate.
        /// </summary>
        public ContentControls ContentControls
        {
            get
            {
                if (_contentControls == null)
                {
                    _contentControls = new ContentControls(InternalObject.GetType().InvokeMember("ContentControls", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _contentControls;
            }
        }

        /// <summary>
        /// Gets the full name of the document.
        /// </summary>
        public string FullName
        {
            get
            {
                return InternalObject.GetType().InvokeMember("FullName", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null) as string;
            }
        }

        /// <summary>
        /// Gets the name of the document.
        /// </summary>
        public string Name
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null) as string;
            }
        }

        /// <summary>
        /// Gets the path of the document.
        /// </summary>
        public string Path
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Path", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null) as string;
            }
        }

        public void PrintOut(bool? Background = null, bool? Append = null, PrintOutRange? Range = null, string OutputFileName = null, int? From = null, int? To = null, PrintOutItem? Item = null,
            int? Copies = null, string Pages = null, PrintOutPages? PageType = null, bool? PrintToFile = null, bool? Collate = null)
        {
            InternalObject.GetType().InvokeMember("PrintOut", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] 
            {
                Background ?? (object)System.Reflection.Missing.Value,
                Append ?? (object)System.Reflection.Missing.Value,
                Range ?? (object)System.Reflection.Missing.Value,
                OutputFileName ?? (object)System.Reflection.Missing.Value,
                From ?? (object)System.Reflection.Missing.Value,
                To ?? (object)System.Reflection.Missing.Value,
                Item ?? (object)System.Reflection.Missing.Value,
                Copies ?? (object)System.Reflection.Missing.Value,
                Pages ?? (object)System.Reflection.Missing.Value,
                PageType ?? (object)System.Reflection.Missing.Value,
                PrintToFile ?? (object)System.Reflection.Missing.Value,
                Collate ?? (object)System.Reflection.Missing.Value
            });
        }

        /// <summary>
        /// Saves a document as PDF or XPS format. (Available from Word 2007)
        /// </summary>
        /// <param name="OutputfileName">The path and file name name of the new PDF or XPS file.</param>
        /// <param name="ExportFormat">Specifies either PDF or XPS format.</param>
        /// <param name="OpenAfterExport">Opens the new file after exporting the contents.</param>
        /// <param name="OptimizeFor">Specifies whether to optimize for screen or print.</param>
        /// <param name="Range">Specifies whether the export range is the entire document, the current page, a range of text, or the current selection. the default is to export the entire document.</param>
        /// <param name="From">Specifies the starting page number, if the Range parameter is set to ExportFromTo.</param>
        /// <param name="To">Specifies the ending page number, if the Range parameter is set to ExportFromTo.</param>
        /// <param name="Item">Specifies whether the export process includes text only or includes text with markup.</param>
        /// <param name="IncludeDocProps">Specifies whether to include document properties in the newly exported file.</param>
        /// <param name="KeepIRM">Specifies whether to copy IRM permissions to an XPS document if the source document has IRM protections. Default value is True.</param>
        /// <param name="CreateBookmarks">Specifies whether to export bookmarks and the type of bookmarks to export.</param>
        /// <param name="DocStructureTags">Specifies whether to include extra data to help screen readers, for example information about the flow and logical organization of the content. Default value is True.</param>
        /// <param name="BitmapMissingFonts">Specifies whether to include a bitmap of the text. Set this parameter to True when font licenses do not permit a font to be embedded in the PDF file. If False, the font is referenced, and the viewer's computer substitutes an appropriate font if the authored one is not available. Default value is True.</param>
        /// <param name="UseISO19005_1">Specifies whether to limit PDF usage to the PDF subset standardized as ISO 19005-1. If True, the resulting files are more reliably self-contained but may be larger or show more visual artifacts due to the restrictions of the format. Default value is False.</param>
        public void ExportAsFixedFormat(string OutputfileName, ExportFormat ExportFormat, bool? OpenAfterExport = false, ExportOptimizeFor? OptimizeFor = ExportOptimizeFor.Print, ExportRange? Range = ExportRange.AllDocument, long? From = null, long? To = null, ExportItem? Item = ExportItem.Content, bool? IncludeDocProps = true, bool? KeepIRM = true, ExportCreateBookmarks? CreateBookmarks = ExportCreateBookmarks.CreateNoBookmarks, bool? DocStructureTags = true, bool? BitmapMissingFonts = true, bool? UseISO19005_1 = false)
        {
            if (String.IsNullOrEmpty(OutputfileName))
                throw new Exception("Output filename may not be blank");

            InternalObject.GetType().InvokeMember("ExportAsFixedFormat", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[]
            {
                OutputfileName,
                ExportFormat,
                OpenAfterExport ?? (object)System.Reflection.Missing.Value,
                OptimizeFor ?? (object)System.Reflection.Missing.Value,
                Range ?? (object)System.Reflection.Missing.Value,
                From ?? (object)System.Reflection.Missing.Value,
                To ?? (object)System.Reflection.Missing.Value,
                Item ?? (object)System.Reflection.Missing.Value,
                IncludeDocProps ?? (object)System.Reflection.Missing.Value,
                KeepIRM ?? (object)System.Reflection.Missing.Value,
                CreateBookmarks ?? (object)System.Reflection.Missing.Value,
                DocStructureTags ?? (object)System.Reflection.Missing.Value,
                BitmapMissingFonts ?? (object)System.Reflection.Missing.Value,
                UseISO19005_1 ?? (object)System.Reflection.Missing.Value,
                (object)System.Reflection.Missing.Value
            });
        }

        /// <summary>
        /// Returns a Range object by using the specified starting and ending character positions.\
        /// Must be manually disposed.
        /// </summary>
        /// <param name="start">The starting character position.</param>
        /// <param name="end">The ending character position.</param>
        /// <returns>Range</returns>
        public Range Range(long? start, long? end)
        {
            var args = new List<object>();
            if (start.HasValue)
                args.Add(start.Value);
            else
                args.Add(System.Reflection.Missing.Value);
            if (end.HasValue)
                args.Add(end.Value);
            else
                args.Add(System.Reflection.Missing.Value);

            return new Range(InternalObject.GetType().InvokeMember("Range", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, args.ToArray()));
        }

        /// <summary>
        /// Gets a boolean representing if the document is read-only or not.
        /// </summary>
        public bool ReadOnly
        {
            get
            {
                return Convert.ToBoolean(InternalObject.GetType().InvokeMember("ReadOnly", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
            }
        }

        /// <summary>
        /// Saves the document.
        /// </summary>
        public void Save()
        {
            InternalObject.GetType().InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// Saves the document under a new filename.
        /// </summary>
        /// <param name="filename">The filename under which to save the document</param>
        public void SaveAs(string filename)
        {
            SaveAs(filename, Enums.SaveFormat.Document);
        }

        /// <summary>
        /// Saves the document under a new filename.
        /// </summary>
        /// <param name="filename">The filename under which to save the document</param>
        /// <param name="fileFormat">The format in which to save the document</param>
        public void SaveAs(string filename, Enums.SaveFormat fileFormat)
        {
            var missing = System.Reflection.Missing.Value;

            InternalObject.GetType().InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { filename, fileFormat });
        }

        /// <summary>
        /// Selects the contents of the specified document.
        /// </summary>
        public void Select()
        {
            InternalObject.GetType().InvokeMember("Select", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }

        /// <summary>
        /// A collection of Section objects in a selection, range, or document. This object is internally cached and does not need to be manually disposed.
        /// </summary>
        public Sections Sections
        {
            get
            {
                if (_sections == null)
                {
                    _sections = new Sections(InternalObject.GetType().InvokeMember("Sections", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _sections;
            }
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_sections != null && !_sections.IsDisposed)
                {
                    _sections.Dispose();
                    _sections = null;
                }

                if (_bookmarks != null && !_bookmarks.IsDisposed)
                {
                    _bookmarks.Dispose();
                    _bookmarks = null;
                }

                if (_contentControls != null && !_contentControls.IsDisposed)
                {
                    _contentControls.Dispose();
                    _contentControls = null;
                }
            }

            base.Dispose(true);
        }

        #endregion
    }
}
