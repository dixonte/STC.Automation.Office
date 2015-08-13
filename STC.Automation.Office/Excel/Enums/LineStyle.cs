using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Excel.Enums
{
    /// <summary>
    /// Specifies the line style for the border.
    /// </summary>
    public enum LineStyle
    {
        /// <summary>
        /// Continuous line.
        /// </summary>
        Continuous = 1,
        /// <summary>
        /// Dashed line.
        /// </summary>
        Dash = -4115,
        /// <summary>
        /// Alternating dashes and dots.
        /// </summary>
        DashDot = 4,
        /// <summary>
        /// Dash followed by two dots.
        /// </summary>
        DashDotDot = 5,
        /// <summary>
        /// Dotted line.
        /// </summary>
        Dot = -4118,
        /// <summary>
        /// Double line.
        /// </summary>
        Double = -4119,
        /// <summary>
        /// No line.
        /// </summary>
        LineStyleNone = -4142,
        /// <summary>
        /// Slanted dashes.
        /// </summary>
        SlantDashDot = 13,

        /// <summary>
        /// Automatic. Not actually part of the Object Model Enum 'xlLineStyle'.
        /// </summary>
        Automatic = -4105,
        /// <summary>
        /// 25% Gray. Not actually part of the Object Model Enum 'xlLineStyle'.
        /// </summary>
        Gray25 = -4124,
        /// <summary>
        /// 50% Gray. Not actually part of the Object Model Enum 'xlLineStyle'.
        /// </summary>
        Gray50 = -4125,
        /// <summary>
        /// 75% Gray. Not actually part of the Object Model Enum 'xlLineStyle'.
        /// </summary>
        Gray75 = -4126
    }
}
