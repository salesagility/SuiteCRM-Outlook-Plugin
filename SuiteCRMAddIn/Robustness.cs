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
            try
            {
                action();
            }
            catch (Exception problem)
            {
                log.Error("Caught top-level error", problem);
            }
        }
    }
}
