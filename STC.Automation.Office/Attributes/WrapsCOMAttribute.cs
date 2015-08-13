using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace STC.Automation.Office.Attributes
{
    [ComVisible(false)]
    [AttributeUsage(AttributeTargets.Class, Inherited = false)]
    internal sealed class WrapsCOMAttribute : Attribute
    {
        private string _progId;
        private Guid[] _mustSupport;

        public WrapsCOMAttribute(string progId, string[] mustSupport)
        {
            _progId = progId;
            _mustSupport = new Guid[mustSupport.Length];
            for (int x = 0; x < mustSupport.Length; x++)
            {
                _mustSupport[x] = new Guid(mustSupport[x]);
            }
        }

        public WrapsCOMAttribute(string progId, string mustSupport)
            : this(progId, new string[] { mustSupport })
        {
        }

        public WrapsCOMAttribute(string progId)
            : this(progId, new string[0])
        {
        }

        public string ProgID
        {
            get
            {
                return _progId;
            }
        }

        public Guid[] MustSupport
        {
            get
            {
                return _mustSupport;
            }
        }
    }
}
