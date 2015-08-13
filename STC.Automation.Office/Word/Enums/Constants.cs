using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Word.Enums
{
    /// <summary>
    /// This enumeration groups together constants used with various Microsoft Word methods.
    /// </summary>
    public enum Constants : long
    {
        /// <summary>
        /// Represents an undefined value.
        /// </summary>
        Undefined = 0x0098967f,
        /// <summary>
        /// Toggles a property's value.
        /// </summary>
        Toggle = 0x0098967e,
        /// <summary>
        /// Indicates that selection will be extended forward using the MoveStartUntil or MoveStartWhile method of the Range or Selection object.
        /// </summary>
        Forward = 0x3fffffff,
        /// <summary>
        /// Indicates that selection will be extended backward using the MoveStartUntil or MoveStartWhile method of the Range or Selection object.
        /// </summary>
        Backward = 0xc0000001,
        /// <summary>
        /// Represents the Auto value for the specified setting.
        /// </summary>
        AutoPosition = 0,
        /// <summary>
        /// Represents the first item in a collection.
        /// </summary>
        First = 1,
        /// <summary>
        /// Represents the creator code for objects created by Microsoft Word. 
        /// </summary>
        CreatorCode = 0x4d535744
    }
}
