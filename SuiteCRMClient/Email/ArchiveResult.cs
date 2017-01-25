using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuiteCRMClient.Email
{
    public class ArchiveResult
    {
        public static ArchiveResult Success(string emailId, IEnumerable<System.Exception> warnings)
        {
            return new ArchiveResult
            {
                EmailId = emailId,
                Problems = warnings,
            };
        }

        public static ArchiveResult Failure(params System.Exception[] exceptions)
        {
            return new ArchiveResult
            {
                Problems = exceptions,
            };
        }

        public string EmailId { get; set; }

        public IEnumerable<System.Exception> Problems { get; set; }

        public bool IsSuccess => !string.IsNullOrEmpty(EmailId);

        public bool IsFailure => !IsSuccess;
    }
}
