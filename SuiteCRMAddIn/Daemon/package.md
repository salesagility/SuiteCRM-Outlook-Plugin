# The Daemon package

The Daemon package is intended to handle background tasks, unloading them from the user interface thread to improve the perceived responsiveness of the application.

## DaemonWorker

Is a singleton class which maintains a queue of tasks ('Actions') and executes them in turn. 
