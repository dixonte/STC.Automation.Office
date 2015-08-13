using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    /// <summary>
    /// Specifies the format to use when saving a document.
    /// </summary>
    public enum SaveFormat
    {
        /// <summary>
        /// Microsoft Office Word format.
        /// </summary>
        Document = 0,
        /// <summary>
        /// Microsoft Word 97 document format.
        /// </summary>
        Document97 = 0,
        /// <summary>
        /// Word template format.
        /// </summary>
        Template = 1,
        /// <summary>
        /// Word 97 template format.
        /// </summary>
        Template97 = 1,
        /// <summary>
        /// Microsoft Windows text format.
        /// </summary>
        Text = 2,
        /// <summary>
        /// Windows text format with line breaks preserved.
        /// </summary>
        TextLineBreaks = 3,
        /// <summary>
        /// Microsoft DOS text format.
        /// </summary>
        DOSText = 4,
        /// <summary>
        /// Microsoft DOS text with line breaks preserved.
        /// </summary>
        DOSTextLineBreaks = 5,
        /// <summary>
        /// Rich text format (RTF).
        /// </summary>
        RTF = 6,
        /// <summary>
        /// Encoded text format.
        /// </summary>
        UnicodeText = 7,
        /// <summary>
        /// Unicode text format.
        /// </summary>
        EncodedText = 7,
        /// <summary>
        /// Standard HTML format.
        /// </summary>
        HTML = 8,
        /// <summary>
        /// Web archive format.
        /// </summary>
        WebArchive = 9,
        /// <summary>
        /// Filtered HTML format.
        /// </summary>
        FilteredHTML = 10,
        /// <summary>
        /// Extensible Markup Language (XML) format.
        /// </summary>
        XML = 11,
        /// <summary>
        /// XML document format.
        /// </summary>
        XMLDocument = 12,
        /// <summary>
        /// XML document format with macros enabled.
        /// </summary>
        XMLDocumentMacroEnabled = 13,
        /// <summary>
        /// XML template format.
        /// </summary>
        XMLTemplate = 14,
        /// <summary>
        /// XML template format with macros enabled.
        /// </summary>
        XMLTemplateMacroEnabled = 15,
        /// <summary>
        /// Word default document file format. For Microsoft Office Word 2007, this is the DOCX format.
        /// </summary>
        DocumentDefault = 16,
        /// <summary>
        /// PDF format.
        /// </summary>
        PDF = 17,
        /// <summary>
        /// XPS format.
        /// </summary>
        XPS = 18,
        /// <summary>
        /// Undocumented.
        /// </summary>
        FlatXML = 19,
        /// <summary>
        /// Undocumented.
        /// </summary>
        FlatXMLMacroEnabled = 20,
        /// <summary>
        /// Undocumented.
        /// </summary>
        FlatXMLTemplate = 21,
        /// <summary>
        /// Undocumented.
        /// </summary>
        FlatXMLTemplateMacroEnabled = 22,
        /// <summary>
        /// Undocumented.
        /// </summary>
        OpenDocumentText = 23
    }
}
