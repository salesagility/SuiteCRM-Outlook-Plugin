using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuiteCRMClient.Exceptions
{
    public class CrmSaveDataException: Exception
    {
        public CrmSaveDataException(string message, Exception inner = null)
            : base(message, inner)
        {
        }
    }
}
