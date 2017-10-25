using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuiteCRMAddIn.Exceptions
{
    /// <summary>
    /// An exception thrown if an action failed but may be retried.
    /// </summary>
    public class ActionFailedException : Exception
    {
        public ActionFailedException(string message) : base(message)
        {

        }

        public ActionFailedException(string message, Exception cause) : base( message, cause)
        {

        }
    }
}
