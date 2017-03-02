using System;
using SuiteCRMClient.Logging;

namespace SuiteCRMAddIn
{
    /// <summary>
    /// Utility class. All event-handlers (in fact: all top-level entry-points) should
    /// have error handling round them. SO this is some boilerplate error-handling.
    /// </summary>
    public class Robustness
    {
        public static void DoOrLogError(ILogger log, Action action)
        {
            DoOrLogError(log, action, "Caught top-level error");
        }

        public static void DoOrLogError(ILogger log, Action action, string message)
        {
            try
            {
                action();
            }
            catch (Exception problem)
            {
                log.Error(message, problem);
            }
        }
    }
}
