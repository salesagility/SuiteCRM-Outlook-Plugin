/// <summary>
/// The Daemon package is intended to handle background tasks, unloading them from the user 
/// interface thread to improve the perceived responsiveness of the application.
/// </summary>
/// <remarks>
/// <para>
/// A singleton instance of DaemonWorker runs as a RepeatingProcess in a thread; it maintains a 
/// queue of DaemonActions, which it executes in turn.
/// </para>
/// <para>
/// Instances of classes implementing DaemonAction are intended to be run essentially once, but may be allowed a number of attempts 
/// (intended to be limited) in case, for example due to network problems, the first attempt(s) fail. 
/// However, DaemonAction is not intended for things which are to be run repeatedly. For that, 
/// specialise [RepeatingProcess](class_suite_c_r_m_add_in_1_1_business_logic_1_1_repeating_process.html).
/// </remarks>
namespace SuiteCRMAddIn.Daemon
{
}
