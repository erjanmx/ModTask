using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TaskScheduler;
using CommandLine;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ModTask
{
    internal class Program
    {

        static List<IRegisteredTask> ListTasks(List<IRegisteredTask> list, ITaskFolder folder)
          {
            foreach (IRegisteredTask task in folder.GetTasks(1))
            {
                list.Add(task);
                
            }

            foreach (ITaskFolder subFolder in folder.GetFolders(1))
            {
               
                ListTasks(list, subFolder);
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(folder);
            return list;

        }
        static IRegisteredTask GetModTask(List<IRegisteredTask> taskList, string modTaskName)
        {
            bool status = false;
            IRegisteredTask foundTask = null;
            try
            {
                foreach (IRegisteredTask task in taskList)
                {
                    if (String.Compare(task.Name, modTaskName) == 0)
                    {
                        Console.WriteLine("[+]Found Requested Task: {0}", task.Name);
                        status = true;
                        foundTask = task;

                    }

                }
                if (!status)
                {
                    Console.WriteLine("[+]Requested Task Was Not Found For Modification");
                    Environment.Exit(0);
                }

                return foundTask;

                

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return null;
            }


        }
        static ITaskService InitTaskScheduler(string serverName, string username, string domain, string password)
        {
            try
            {
                ITaskService ts = new TaskScheduler.TaskScheduler();
                if((!String.IsNullOrEmpty(username) || !String.IsNullOrEmpty(password) || !String.IsNullOrEmpty(domain)) && String.IsNullOrEmpty(serverName))
                {
                    Console.WriteLine("[+] Username,Password and Domain settings only for remote use, if you want local execution, rerun without those flags");
                    Environment.Exit(0);
                }
                if (!String.IsNullOrEmpty(username) && !String.IsNullOrEmpty(password) && !String.IsNullOrEmpty(domain) && !String.IsNullOrEmpty(serverName))
                {
                    
                    ts.Connect(serverName, username, domain, password);

                }
                if (!String.IsNullOrEmpty(username) && !String.IsNullOrEmpty(password) && !String.IsNullOrEmpty(serverName))
                {
                    
                    ts.Connect(serverName, username, "", password);

                }
                if (String.IsNullOrEmpty(username) && String.IsNullOrEmpty(password) && String.IsNullOrEmpty(domain) && String.IsNullOrEmpty(serverName)) { 
                   
                    ts.Connect();
                }

                
                
                
                ITaskFolder rootFolder = ts.GetFolder(@"\");
                return ts;
            }
            catch(Exception e)
            {
                Console.WriteLine(e.ToString());
                return null;

            }

           

        }

      static void ListTaskStart(string serverName,string username,string domain, string password, bool acl)
        {
            try
            {
                ITaskService ts = InitTaskScheduler(serverName, username, domain, password);
                ITaskFolder rootFolder = ts.GetFolder(@"\");
                if (rootFolder == null)
                {
                    Environment.Exit(0);
                }
                List<IRegisteredTask> tasks = new List<IRegisteredTask>();
                ListTasks(tasks, rootFolder);
                string result = ""
                ;
                foreach (IRegisteredTask task in tasks)
                {
                   Console.WriteLine("Task Name: "+task.Name);
                    if (!String.IsNullOrEmpty(task.Definition.Principal.UserId))
                    {
                        Console.WriteLine("Run As Principal: {0}", task.Definition.Principal.UserId);
                    }
                    if (!String.IsNullOrEmpty(task.Definition.Principal.GroupId))
                    {
                        Console.WriteLine("Run As Principal: {0}", task.Definition.Principal.GroupId);
                    }
                    Console.WriteLine("Last RunTime: {0}",task.LastRunTime);
                    if ((int)task.State == 4)
                    {
                        Console.WriteLine("Status: Running");
                        
                    }
                    if (acl)
                    {
                        result = task.GetSecurityDescriptor(1|2|4);
                        Console.WriteLine("SDDL: "+result+"\n");
                        
                    }
                    Console.WriteLine("----------------------------");
                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            }
       static void ModTaskStart(string serverName,string username,string domain,string password, string taskToMod, string argExe,string exeArgs,bool sys,bool com, string classID,bool boottrigger, bool timetrigger,string repDur,string repInter)
       {
            try
            {
                IRunningTask runTask = null;
                Boolean demandChanged = false;
                ITaskService ts = InitTaskScheduler(serverName, username, domain, password);
                ITaskFolder rootFolder = ts.GetFolder(@"\");
                List<IRegisteredTask> tasks = new List<IRegisteredTask>();
                IRegisteredTask returnedTask = null;
                ListTasks(tasks, rootFolder);
                returnedTask = GetModTask(tasks, taskToMod);
                Console.WriteLine("[+]Task Path: {0}", returnedTask.Path);
                ITaskDefinition taskdef = returnedTask.Definition;
                IActionCollection taskActionCol = taskdef.Actions;
                taskActionCol.Clear();
                if (com)
                {
                    IAction comTaskAction = taskActionCol.Create(_TASK_ACTION_TYPE.TASK_ACTION_COM_HANDLER);
                    IComHandlerAction comAction = (IComHandlerAction)comTaskAction;
                    comAction.ClassId = classID;
                   
                }
                else
                {
                    IAction taskAction = taskActionCol.Create(_TASK_ACTION_TYPE.TASK_ACTION_EXEC);
                    IExecAction execAction = (IExecAction)taskAction;
                    execAction.Path = argExe;
                    execAction.Arguments = exeArgs;
                }
                ITriggerCollection trigCollection = taskdef.Triggers;
                trigCollection.Clear();
                if (timetrigger)
                {
                    
                    ITrigger trig = trigCollection.Create(_TASK_TRIGGER_TYPE2.TASK_TRIGGER_DAILY);
                    IDailyTrigger dailyTrig = (IDailyTrigger)trig;
                    dailyTrig.StartBoundary = DateTime.UtcNow.AddMinutes(1.0).ToString("o");
                    dailyTrig.Repetition.Interval = repInter;
                    dailyTrig.Repetition.Duration = repDur;
                    

                }
                if (boottrigger)
                {
                    ITrigger trig2 = trigCollection.Create(_TASK_TRIGGER_TYPE2.TASK_TRIGGER_BOOT);
                    IBootTrigger boot = (IBootTrigger)trig2;
                    boot.Enabled = true;
                    
                }
               
               
                string path = returnedTask.Path.Substring(0,returnedTask.Path.LastIndexOf("\\"));
                
                ITaskSettings tasksettings = taskdef.Settings;
                
                if((int)returnedTask.State == 4)
                {
                    Console.WriteLine("[+]Task Is Currently Running, Stopping The Task Before Modification");
                    returnedTask.Stop(0);
                    Console.WriteLine("[+]Stopped Task Successfully");
                  
                }
                if (tasksettings.AllowDemandStart != true)
                {
                    tasksettings.AllowDemandStart = true;
                    Console.WriteLine("[+]Enabled AllowDemandStart setting");
                    demandChanged = true;
                }

                
                ITaskFolder returnedFolder = ts.GetFolder(path);
                
                returnedFolder.RegisterTaskDefinition(returnedTask.Name, taskdef, 4, null, null, _TASK_LOGON_TYPE.TASK_LOGON_NONE, null);
                if (sys)
                {
                    
                    runTask = returnedTask.RunEx(null, 0, 0, "NT AUTHORITY\\SYSTEM");
                   
                    Console.WriteLine("[+]Successfully Ran Task: {0}",returnedTask.Name);
                   

                }
                else
                {
                
                    runTask = returnedTask.Run(null);
                   
                    Console.WriteLine("[+]Successfully Ran Task: {0}",returnedTask.Name);
                    
                }
                if (timetrigger == false || boottrigger == true)
                {
                    if (demandChanged == true)
                    {
                        tasksettings.AllowDemandStart = false;
                    }

                    returnedFolder.RegisterTaskDefinition(returnedTask.Name, returnedTask.Definition, 4, null, null, _TASK_LOGON_TYPE.TASK_LOGON_NONE, null);
                    Console.WriteLine("[+]Successfully Reverted Task: {0}", returnedTask.Name);
                }



                foreach (IRegisteredTask task in tasks)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(task);
                }


            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
        static void SelectTask(string serverName, string username, string domain, string password, string selectTask)
        {
           ITaskService ts = InitTaskScheduler(serverName, username, domain, password);
           ITaskFolder rootFolder = ts.GetFolder(@"\");
           List<IRegisteredTask> tasks = new List<IRegisteredTask>();
           IRegisteredTask returnedTask = null;
           ListTasks(tasks, rootFolder);
           returnedTask = GetModTask(tasks, selectTask);
           Console.WriteLine(returnedTask.Xml);
        }
       
        static void Logo()
        {
            string logo = @"

         .---.  ,'|""\   _______  .--.     .---. ,-. .-. 
|\    /|/ .-. ) | |\ \ |__   __|/ /\ \   ( .-._)| |/ /  
|(\  / || | |(_)| | \ \  )| |  / /__\ \ (_) \   | | /   
(_)\/  || | | | | |  \ \(_) |  |  __  | _  \ \  | | \   
| \  / |\ `-' / /(|`-' /  | |  | |  |)|( `-'  ) | |) \  
| |\/| | )---' (__)`--'   `-'  |_|  (_) `----'  |((_)-' 
'-'  '-'(_)                                     (_)     

";
            Console.WriteLine(logo);
        }
        class ArgOpts
        {
            [Option('m',"mode",Required =true,HelpText = "Options: \n list - Lists Tasks on the targeted system \n          [+] ModTask.exe list --servername [SERVER] --username [USER] --password [PASSWORD] --domain [DOMAIN] [--sddl] \n modify - Takes a supplied Task and modifies it for binary execution \n          [+] Modtask.exe modify --taskName [TASK] --servername [SERVER] --username [USER] --password [PASSWORD] --domain [DOMAIN] --exePath [EXE] --exeArgs [EXEARGS] [--sys] [--com] --comClassID [CLASSID] [--timetrigger] --repDuration [PT1M] --repInterval [PT5M] [--boottrigger]\n select - Provides detailed XML information on a supplied Task \n          [+] ModTask.exe select --taskName [TASK] --servername [SERVER] --username [USER] --password [PASSWORD] --domain [DOMAIN]")]
            public string Mode { get; set; }

            [Option('t', "taskName",HelpText ="Task to Modify - To be Used with modify/select mode")]
            public String taskName { get; set; }

            [Option('s', "servername",HelpText ="Server to connect to")]
            public String servername { get; set; }

            [Option('u', "username",HelpText ="Username to Connect With - Not including this option will cause the application to run as the current user - Remote Only")]
            public String username { get; set; }

            [Option('p',"password", HelpText = "Password to Connect With - Not including this option will cause the application to run as the current user - Remote Only")]
            public String password { get; set; }
            [Option('d', "domain", HelpText = "Domain to Connect To - Not including this option will cause the application to run as the current user - Remote Only")]
            public String domain { get; set; }

            [Option('l', "sys",HelpText = "Run Task as SYSTEM - To be used with modify mode - Optional")]
            public Boolean sys { get; set; }

            [Option('e', "exePath",HelpText ="Exe to Run with selected Task")]
            public string exePath { get; set; }

            [Option('a', "exeArgs",HelpText = "Exe Args to Run with selected Task")]
            public string exeArgs { get; set; }

            [Option('c', "sddl", HelpText = "Displays sddl strings for listed tasks - To be used with list mode - Optional")]
            public Boolean acl { get; set; }

            [Option('t', "timetrigger", HelpText = "Enables Repition Patterns to be added to specified Task - To be used with modify mode - Manual Cleanup required - Optional")]
            public Boolean timetrigger { get; set; }

            [Option('b', "boottrigger", HelpText = "Enables Boot Trigger to be added to specified Task - To be used with modify mode - Manual Cleanup required - Optional")]
            public Boolean boottrigger { get; set; }


            [Option('r', "repInterval", HelpText = "Time Interval to run the specified Task - Ex. PT5M - 5 MINUTES, PT5H - 5 HOURS - To be used with timetrigger option")]
            public string repInter { get; set; }

            [Option('d', "repDuration", HelpText = "Time Duration to run the specified Task - Ex. PT5M - 5 MINUTES, PT5H - 5 HOURS - To be used with timetrigger option")]
            public string repDur { get; set; }

            [Option('o', "com", HelpText = "Enables a COM Object for specifed Task- Optional")]
            public Boolean com { get; set; }

            [Option('i', "comClassID", HelpText = "COM Object Class ID to run with selected Task Ex. {00000000-0000-0000-0000-000000000000} -- To be used with com option")]
            public string classID { get; set; }

           




        }
        static void Main(string[] args)
        {
            
            string taskToMod = "";
            string serverName = "";
            string username = "";
            string password = "";
            string domain = "";
            string argExe = "";
            string exeArgs = "";
            string classID = "";
            string repInter = "";
            string repDur = "";
            bool sys = false;
            bool sddl = false;
            bool com = false;
            bool timetrigger = false;
            bool boottrigger = false;
            
           
            Logo();
            Parser.Default.ParseArguments<ArgOpts>(args).WithParsed(options =>
            {
                
                
                if(options.Mode == "list")
                {
                    if (!String.IsNullOrEmpty(options.taskName))
                    {
                        Console.WriteLine("[+] Task Name should only be used in modify mode or select mode");
                        Environment.Exit(0);
                    }
                    if (!String.IsNullOrEmpty((options.servername)))
                    {
                        serverName = options.servername;
                    }
                    if (!String.IsNullOrEmpty((options.username)))
                    {
                        username = options.username;
                    }
                    if (!String.IsNullOrEmpty((options.password)))
                    {
                        password = options.password;
                    }
                    if (!String.IsNullOrEmpty((options.domain)))
                    {
                        domain = options.domain;
                    }
                    if (options.acl == true){
                        sddl = options.acl;
                    }

                    ListTaskStart(serverName, username, domain, password, sddl);
                    Environment.Exit(1);


                }
             if(options.Mode == "modify")
                {
                    if (!String.IsNullOrEmpty((options.servername)))
                    {
                        serverName = options.servername;
                    }
                    if (!String.IsNullOrEmpty((options.username)))
                    {
                        username = options.username;
                    }
                    if (!String.IsNullOrEmpty((options.password)))
                    {
                        password = options.password;
                    }
                    if (!String.IsNullOrEmpty((options.domain)))
                    {
                        domain = options.domain;
                    }
                    if (options.timetrigger)
                    {
                        timetrigger = options.timetrigger;
                        repDur = options.repDur;
                        repInter = options.repInter;

                    }
                    if (options.boottrigger)
                    {
                        boottrigger = options.boottrigger;
                    }
                    if(options.com)
                    {
                        com = options.com;
                        classID = options.classID;
                        if (!String.IsNullOrEmpty(options.exePath))
                        {
                            Console.WriteLine("[+] Exe can't be specified if using COM objects");
                            Environment.Exit(0);
                        }
                        if (!String.IsNullOrEmpty(options.exeArgs))
                        {
                            Console.WriteLine("[+] Exe Arguments can't be specified if using COM objects");
                            Environment.Exit(0);

                        }
                    }
                    if (!String.IsNullOrEmpty(options.exePath))
                    {
                        argExe = options.exePath;
                    }
                    if (!String.IsNullOrEmpty(options.exeArgs))
                    {
                        exeArgs = options.exeArgs;
                    }
                    if (!String.IsNullOrEmpty((options.taskName)))
                    {
                        taskToMod = options.taskName;
                    }
                    else
                    {
                        Console.WriteLine("[+] TaskName to modify is required for this mode");
                        Environment.Exit(0);
                    }
                    if(options.sys == true)
                    {
                        sys = true;
                    }

                    ModTaskStart(serverName, username, domain, password, taskToMod, argExe, exeArgs, sys, com,classID,boottrigger,timetrigger,repDur,repInter);
                    Environment.Exit(1);


                }
                if (options.Mode == "select")
                {
                    if (!String.IsNullOrEmpty((options.servername)))
                    {
                        serverName = options.servername;
                    }
                    if (!String.IsNullOrEmpty((options.username)))
                    {
                        username = options.username;
                    }
                    if (!String.IsNullOrEmpty((options.password)))
                    {
                        password = options.password;
                    }
                    if (!String.IsNullOrEmpty((options.domain)))
                    {
                        domain = options.domain;
                    }
                    if (!String.IsNullOrEmpty((options.taskName)))
                    {
                        taskToMod = options.taskName;
                    }
                    SelectTask(serverName, username, domain, password, taskToMod);
                    Environment.Exit(1);


                }
                {

                }

            });
            
           
        }
    }
}
