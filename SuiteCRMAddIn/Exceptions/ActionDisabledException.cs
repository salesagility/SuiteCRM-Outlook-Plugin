namespace SuiteCRMAddIn.Exceptions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// A mechanism for temporarily disabling an action.
    /// </summary>
    public class ActionDisabledException : ActionFailedException
    {
        public ActionDisabledException() : base("This action is dissabled") { }
    }
}
