namespace SuiteCRMAddIn.BusinessLogic
{
    public interface DaemonAction
    {
        /// <summary>
        /// Get a description of this action.
        /// </summary>
        string Description {
            get;
        }

        /// <summary>
        /// Perform this action.
        /// </summary>
        void Perform();
    }
}
