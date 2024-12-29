# ModTask - Task Scheduler Attack Tool

ModTask is a C# Red Teaming Attack Tool that can be utilized for:

- Listing Scheduled Tasks along with their SDDL strings and key information, locally and remotely.
- Selecting a specific Scheduled Task for a detailed overview of its configuration settings.
- Modifying a Scheduled Task, locally and remotely. Utilizing either an Exe file path and arguments or a COM object Class ID for execution. Useful for lateral movement and persistence scenarios.
- Supports mutiple trigger modifications such as Startup Boot Triggers and Daily Triggers with Repetition Patterns.
- Built in cleanup functionality to revert the task to its original state before any modifications had taken place.
  
![pic](https://github.com/user-attachments/assets/5ffc7560-609c-48bc-92b9-c0852e369086)

*Does require Administrative access on the remote or local machine in order to work.*
