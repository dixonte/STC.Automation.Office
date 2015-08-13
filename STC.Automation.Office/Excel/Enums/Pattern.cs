using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Specifies the parameter on which the data should be sorted.
    /// </summary>
    public enum Pattern
    {
        ///<summary>
        ///Excel controls the pattern.
        ///</summary>
        PatternAutomatic = -4105,

        ///<summary>
        ///Checkerboard.
        ///</summary>
        PatternChecker = 9,

        ///<summary>
        ///Criss-cross lines.
        ///</summary>
        PatternCrissCross = 16,

        ///<summary>
        ///Dark diagonal lines running from the upper left to the lower right.
        ///</summary>
        PatternDown = -4121,

        ///<summary>
        ///16% gray.
        ///</summary>
        PatternGray16 = 17,

        ///<summary>
        ///25% gray.
        ///</summary>
        PatternGray25 = -4124,

        ///<summary>
        ///50% gray.
        ///</summary>
        PatternGray50 = -4125,

        ///<summary>
        ///75% gray.
        ///</summary>
        PatternGray75 = -4126,

        ///<summary>
        ///8% gray.
        ///</summary>
        PatternGray8 = 18,

        ///<summary>
        ///Grid.
        ///</summary>
        PatternGrid = 15,

        ///<summary>
        ///Dark horizontal lines.
        ///</summary>
        PatternHorizontal = -4128,

        ///<summary>
        ///Light diagonal lines running from the upper left to the lower right.
        ///</summary>
        PatternLightDown = 13,

        ///<summary>
        ///Light horizontal lines.
        ///</summary>
        PatternLightHorizontal = 11,

        ///<summary>
        ///Light diagonal lines running from the lower left to the upper right.
        ///</summary>
        PatternLightUp = 14,

        ///<summary>
        ///Light vertical bars.
        ///</summary>
        PatternLightVertical = 12,

        ///<summary>
        ///No pattern.
        ///</summary>
        PatternNone = -4142,

        ///<summary>
        ///75% dark moiré.
        ///</summary>
        PatternSemiGray75 = 10,

        ///<summary>
        ///Solid color.
        ///</summary>
        PatternSolid = 1,

        ///<summary>
        ///Dark diagonal lines running from the lower left to the upper right.
        ///</summary>
        PatternUp = -4162,

        ///<summary>
        ///Dark vertical bars.
        ///</summary>
        PatternVertical = -4166
    }
}
