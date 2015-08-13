using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    /// <summary>
    /// Specifies a Microsoft Office Word field.
    /// Unless otherwise specified, the field types described in this enumeration can be added interactively to a Word document by using the Field dialog box.
    /// See the Word Help for more information about specific field codes.
    /// </summary>
    public enum FieldType
    {
        /// <summary>
        /// Add-in field. Not available through the Field dialog box. Used to store data that is hidden from the user interface.
        /// </summary>
        Addin = 81,
        /// <summary>
        /// AddressBlock field.
        /// </summary>
        AddressBlock = 93,
        /// <summary>
        /// Advance field.
        /// </summary>
        Advance = 84,
        /// <summary>
        /// Ask field.
        /// </summary>
        Ask = 38,
        /// <summary>
        /// Author field.
        /// </summary>
        Author = 17,
        /// <summary>
        /// AutoNum field.
        /// </summary>
        AutoNum = 54,
        /// <summary>
        /// AutoNumLgl field.
        /// </summary>
        AutoNumLegal = 53,
        /// <summary>
        /// AutoNumOut field.
        /// </summary>
        AutoNumOutline = 52,
        /// <summary>
        /// AutoText field.
        /// </summary>
        AutoText = 79,
        /// <summary>
        /// AutoTextList field.
        /// </summary>
        AutoTextList = 89,
        /// <summary>
        /// BarCode field.
        /// </summary>
        BarCode = 63,
        /// <summary>
        /// BidiOutline field.
        /// </summary>
        BidiOutline = 92,
        /// <summary>
        /// Comments field.
        /// </summary>
        Comments = 19,
        /// <summary>
        /// Compare field.
        /// </summary>
        Compare = 80,
        /// <summary>
        /// CreateDate field.
        /// </summary>
        CreateDate = 21,
        /// <summary>
        /// Data field.
        /// </summary>
        Data = 40,
        /// <summary>
        /// Database field.
        /// </summary>
        Database = 78,
        /// <summary>
        /// Date field.
        /// </summary>
        Date = 31,
        /// <summary>
        /// DDE field. No longer available through the Field dialog box, but supported for documents created in earlier versions of Word.
        /// </summary>
        DDE = 45,
        /// <summary>
        /// DDEAuto field. No longer available through the Field dialog box, but supported for documents created in earlier versions of Word.
        /// </summary>
        DDEAuto = 46,
        /// <summary>
        /// DocProperty field.
        /// </summary>
        DocProperty = 85,
        /// <summary>
        /// DocVariable field.
        /// </summary>
        DocVariable = 64,
        /// <summary>
        /// EditTime field.
        /// </summary>
        EditTime = 25,
        /// <summary>
        /// Embedded field.
        /// </summary>
        Embed = 58,
        /// <summary>
        /// Empty field. Acts as a placeholder for field content that has not yet been added. A field added by pressing Ctrl+F9 in the user interface is an Empty field.
        /// </summary>
        Empty = -1,
        /// <summary>
        /// = (Formula) field.
        /// </summary>
        Expression = 34,
        /// <summary>
        /// FileName field.
        /// </summary>
        FileName = 29,
        /// <summary>
        /// FileSize field.
        /// </summary>
        FileSize = 69,
        /// <summary>
        /// Fill-In field.
        /// </summary>
        FillIn = 39,
        /// <summary>
        /// FootnoteRef field. Not available through the Field dialog box. Inserted programmatically or interactively.
        /// </summary>
        FootnoteRef = 5,
        /// <summary>
        /// FormCheckBox field.
        /// </summary>
        FormCheckBox = 71,
        /// <summary>
        /// FormDropDown field.
        /// </summary>
        FormDropDown = 83,
        /// <summary>
        /// FormText field.
        /// </summary>
        FormTextInput = 70,
        /// <summary>
        /// EQ (Equation) field.
        /// </summary>
        Formula = 49,
        /// <summary>
        /// Glossary field. No longer supported in Word.
        /// </summary>
        Glossary = 47,
        /// <summary>
        /// GoToButton field.
        /// </summary>
        GoToButton = 50,
        /// <summary>
        /// GreetingLine field.
        /// </summary>
        GreetingLine = 94,
        /// <summary>
        /// HTMLActiveX field. Not currently supported.
        /// </summary>
        HTMLActiveX = 91,
        /// <summary>
        /// Hyperlink field.
        /// </summary>
        Hyperlink = 88,
        /// <summary>
        /// If field.
        /// </summary>
        If = 7,
        /// <summary>
        /// Import field. Cannot be added through the Field dialog box, but can be added interactively or through code.
        /// </summary>
        Import = 55,
        /// <summary>
        /// Include field. Cannot be added through the Field dialog box, but can be added interactively or through code.
        /// </summary>
        Include = 36,
        /// <summary>
        /// IncludePicture field.
        /// </summary>
        IncludePicture = 67,
        /// <summary>
        /// IncludeText field.
        /// </summary>
        IncludeText = 68,
        /// <summary>
        /// Index field.
        /// </summary>
        Index = 8,
        /// <summary>
        /// XE (Index Entry) field.
        /// </summary>
        IndexEntry = 4,
        /// <summary>
        /// Info field.
        /// </summary>
        Info = 14,
        /// <summary>
        /// Keywords field.
        /// </summary>
        KeyWord = 18,
        /// <summary>
        /// LastSavedBy field.
        /// </summary>
        LastSavedBy = 20,
        /// <summary>
        /// Link field.
        /// </summary>
        Link = 56,
        /// <summary>
        /// ListNum field.
        /// </summary>
        ListNum = 90,
        /// <summary>
        /// MacroButton field.
        /// </summary>
        MacroButton = 51,
        /// <summary>
        /// MergeField field.
        /// </summary>
        MergeField = 59,
        /// <summary>
        /// MergeRec field.
        /// </summary>
        MergeRec = 44,
        /// <summary>
        /// MergeSeq field.
        /// </summary>
        MergeSeq = 75,
        /// <summary>
        /// Next field.
        /// </summary>
        Next = 41,
        /// <summary>
        /// NextIf field.
        /// </summary>
        NextIf = 42,
        /// <summary>
        /// NoteRef field.
        /// </summary>
        NoteRef = 72,
        /// <summary>
        /// NumChars field.
        /// </summary>
        NumChars = 28,
        /// <summary>
        /// NumPages field.
        /// </summary>
        NumPages = 26,
        /// <summary>
        /// NumWords field.
        /// </summary>
        NumWords = 27,
        /// <summary>
        /// OCX field. Cannot be added through the Field dialog box, but can be added through code by using the AddOLEControl method of the Shapes collection or of the InlineShapes collection.
        /// </summary>
        OCX = 87,
        /// <summary>
        /// Page field.
        /// </summary>
        Page = 33,
        /// <summary>
        /// PageRef field.
        /// </summary>
        PageRef = 37,
        /// <summary>
        /// Print field.
        /// </summary>
        Print = 48,
        /// <summary>
        /// PrintDate field.
        /// </summary>
        PrintDate = 23,
        /// <summary>
        /// Private field.
        /// </summary>
        Private = 77,
        /// <summary>
        /// Quote field.
        /// </summary>
        Quote = 35,
        /// <summary>
        /// Ref field.
        /// </summary>
        Ref = 3,
        /// <summary>
        /// RD (Reference Document) field.
        /// </summary>
        RefDoc = 11,
        /// <summary>
        /// RevNum field.
        /// </summary>
        RevisionNum = 24,
        /// <summary>
        /// SaveDate field.
        /// </summary>
        SaveDate = 22,
        /// <summary>
        /// Section field.
        /// </summary>
        Section = 65,
        /// <summary>
        /// SectionPages field.
        /// </summary>
        SectionPages = 66,
        /// <summary>
        /// Seq (Sequence) field.
        /// </summary>
        Sequence = 12,
        /// <summary>
        /// Set field.
        /// </summary>
        Set = 6,
        /// <summary>
        /// Shape field. Automatically created for any drawn picture.
        /// </summary>
        Shape = 95,
        /// <summary>
        /// SkipIf field.
        /// </summary>
        SkipIf = 43,
        /// <summary>
        /// StyleRef field.
        /// </summary>
        StyleRef = 10,
        /// <summary>
        /// Subject field.
        /// </summary>
        Subject = 16,
        /// <summary>
        /// Macintosh only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
        /// </summary>
        Subscriber = 82,
        /// <summary>
        /// Symbol field.
        /// </summary>
        Symbol = 57,
        /// <summary>
        /// Template field.
        /// </summary>
        Template = 30,
        /// <summary>
        /// Time field.
        /// </summary>
        Time = 32,
        /// <summary>
        /// Title field.
        /// </summary>
        Title = 15,
        /// <summary>
        /// TOA (Table of Authorities) field.
        /// </summary>
        TOA = 73,
        /// <summary>
        /// TOA (Table of Authorities Entry) field.
        /// </summary>
        TOAEntry = 74,
        /// <summary>
        /// TOC (Table of Contents) field.
        /// </summary>
        TOC = 13,
        /// <summary>
        /// TOC (Table of Contents Entry) field.
        /// </summary>
        TOCEntry = 9,
        /// <summary>
        /// UserAddress field.
        /// </summary>
        UserAddress = 62,
        /// <summary>
        /// UserInitials field.
        /// </summary>
        UserInitials = 61,
        /// <summary>
        /// UserName field.
        /// </summary>
        UserName = 60,
        /// <summary>
        /// Bibliography field.
        /// </summary>
        Bibliography = 97,
        /// <summary>
        /// Citation field.
        /// </summary>
        Citation = 96
    }
}
