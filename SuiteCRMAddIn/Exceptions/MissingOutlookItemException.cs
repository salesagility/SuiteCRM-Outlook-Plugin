using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuiteCRMAddIn.Exceptions
{
    public class MissingOutlookItemException : Exception
    {
        /// <summary>
        /// Construct a new instance of a MissingOutlookItemException.
        /// </summary>
        /// <param name="entryId">The entry id which was searched for but not found.</param>
        public MissingOutlookItemException(string entryId) : base(
            $"An outlook item with entry ID '{entryId}' could not be found")
        {
            
        }
    }
}
