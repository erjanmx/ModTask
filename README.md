ModTask is a C# Red Teaming Attack Tool that can utilized for:

- Listing Scheduled Tasks along with their SDDL strings and key information, locally and remotely.
- Selecting a specific Scheduled Task for a detailed overview of its configuration settings.
- Modifying a Scheduled Task, locally and remotely. Utilzing either an Exe file path and arguments or a COM object Class ID for execution. Useful for lateral movement and persistence scenarios.
- Supports mutiple trigger modifications such as Startup Boot Triggers and Daily Triggers with Repetition Patterns.
- Buitlin cleanup functionality to revert the task to its orginal state before any modifications had taken place.

Does require Administrative access on the remote or local machine in order to work.
