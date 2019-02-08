using System;
using SuiteCRMAddIn.BusinessLogic;
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

        /// <summary>
        /// Do this action and, if an error occurs, invoke the error handler on it with this message.
        /// </summary>
        /// <remarks>
        /// \todo this method is duplicated in ErrorHandler, and the copy in ErrorHandler is preferred;
        /// in the next release it is intended to remove the Robustness class and move its functionality
        /// to ErrorHandler.
        /// </remarks>
        /// <param name="action">The action to perform</param>
        /// <param name="message">A string describing what the action was intended to achieve.</param>
        public static void DoOrHandleError(Action action, string message)
        {
            try
            {
                action();
            }
            catch (Exception problem)
            {
                ErrorHandler.Handle(message, problem);
            }
        }
    }
}
