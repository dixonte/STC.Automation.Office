using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Core.Enums
{
    public enum TriState : uint
    {
        True = 0xffffffff,
        False = 0,
        CTrue = 1,
        Toggle = 0xfffffffd,
        Mixed = 0xfffffffe
    }
}
