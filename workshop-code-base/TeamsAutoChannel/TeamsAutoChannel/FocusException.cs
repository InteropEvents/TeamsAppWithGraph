using System;

namespace Microsoft.Office.Interop.TeamsAuto
{
    public class FocusException : Exception
    {
        public FocusException(string msg) : base(msg) { }
    }
}
