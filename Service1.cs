using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using Microsoft.Win32;
using System.IO;

namespace Task_Queue
{
    public partial class Service1 : ServiceBase
    {
        private static int currentTaskCount = 0;
        private static readonly object lockObj = new object();
        private static string registryPath = @"Software\Task_Queue\Claims";
        private static string tasksRegistryPath = @"Software\Task_Queue\Tasks";
        private static string logPath = @"C:\Windows\Logs\Task_Queue_18-11-2013.log";
        private static string parametersRegistryPath = @"SOFTWARE\Task_Queue\Parameters";
        private static int Task_Claim_Check_Period, Task_Execution_Duration, Task_Execution_Quantity;

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Task_Claim_Check_Period = ReadRegistry(parametersRegistryPath, "Task_Claim_Check_Period") != null ? Convert.ToInt32(ReadRegistry(parametersRegistryPath, "Task_Claim_Check_Period")) : 30;
            Task_Execution_Duration = ReadRegistry(parametersRegistryPath, "Task_Execution_Duration") != null ? Convert.ToInt32(ReadRegistry(parametersRegistryPath, "Task_Execution_Duration")) : 60;
            Task_Execution_Quantity = ReadRegistry(parametersRegistryPath, "Task_Execution_Quantity") != null ? Convert.ToInt32(ReadRegistry(parametersRegistryPath, "Task_Execution_Quantity")) : 1;

            System.Timers.Timer t = new System.Timers.Timer(Task_Claim_Check_Period * 1000);
            t.Elapsed += new ElapsedEventHandler(CheckerCycle);
            t.Enabled = true;

            System.Timers.Timer t2 = new System.Timers.Timer(1000);
            t2.Elapsed += new ElapsedEventHandler(ExecuteTasks);
            t2.Enabled = true;

            WriteLog("Службу успішно запущено!");
        }

        protected override void OnStop()
        {
            WriteLog("Службу було зупинено!");
        }

        private static bool IsValidTaskName(string taskName)
        {
            return Regex.IsMatch(taskName, @"^Task_\d{4}$");
        }

        private static bool IsTaskServiced(string taskName)
        {
            return Regex.IsMatch(taskName, @"^Task_\d{4}.*$|Task_\d{4}.*-Queued$|Task_\d{4}.*-In progress - \d+.*$|Task_\d{4}-Completed$");
        }

        private static void CheckerCycle(object source, ElapsedEventArgs e)
        {
            var key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(registryPath, true);
            bool same = false;

            string taskToProcess = null;
            int minNumber = int.MaxValue;
            List<string> uniqueTaskIDs = new List<string>();

            foreach (string taskName in key.GetValueNames())
            {
                if (key == null)
                {
                    return;
                }
                foreach (string task in RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(tasksRegistryPath, true).GetValueNames())
                {
                    if (task.Contains(taskName))
                    {
                        same = true;

                    }
                }
                if (same)
                {
                    key.DeleteValue(taskName);
                    WriteLog($"ПОМИЛКА розміщення заявки {taskName}: номер вже існує...");
                    return;
                }
                string pattern = @"^Task_\d{4}$";
                Regex regex = new Regex(pattern);
                if (!regex.IsMatch(taskName))
                {
                    key.DeleteValue(taskName);
                    WriteLog($"ПОМИЛКА розміщення заявки {taskName}: неправильний ситнаксис...");
                    return;

                }
                if (IsValidTaskName(taskName))
                {
                    int number = int.Parse(taskName.Split('_')[1]);
                    string taskID = "Task_" + number.ToString();

                    int duplicateCount = key.GetValueNames().Count(name => name.Contains(taskID));

                    if (duplicateCount > 1)
                    {
                        string duplicateTask = key.GetValueNames().FirstOrDefault(name => name.Contains(taskID) && name != taskName);
                        if (duplicateTask != null && IsTaskServiced(duplicateTask))
                        {
                            key.DeleteValue(taskName);
                            WriteLog($"ПОМИЛКА розміщення заявки {taskID}: номер вже існує...");
                            continue;
                        }
                    }

                    uniqueTaskIDs.Add(taskID);

                    if (number < minNumber)
                    {
                        minNumber = number;
                        taskToProcess = taskName;
                    }
                }
                else if (IsTaskServiced(taskName))
                {
                    continue;
                }
                else
                {
                    key.DeleteValue(taskName);
                    WriteLog($"ПОМИЛКА розміщення заявки {taskName}: неправильний ситнаксис...");
                }
            }

            if (taskToProcess != null)
            {
                WriteLog($"Задача {taskToProcess} успішно прийнята в обробку...");

                using (var tasksKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(tasksRegistryPath, true))
                {
                    tasksKey.SetValue($"Task_{minNumber.ToString("D4")}-Queued", "0000000000000000", RegistryValueKind.String);
                }

                //string newTaskName = "Task_" + minNumber.ToString("D4") + "-[0000000]-Queued";
                key.DeleteValue(taskToProcess); // Видалити завдання з "Software\Task_Queue\Claims"

            }
            key.Close();
        }

        private static void ExecuteTasks(object source, ElapsedEventArgs e)
        {
            var key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(tasksRegistryPath, true);
            var key0 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(registryPath, true);
            if (key == null)
            {
                return;
            }

            var tasks = key.GetValueNames().Where(name => name.Contains("-Queued")).OrderBy(name => GetTaskIdFromTaskName(name)).ToArray();

            Parallel.ForEach(tasks, new ParallelOptions { MaxDegreeOfParallelism = Task_Execution_Quantity }, taskName =>
            {
                lock (lockObj)
                {
                    if (currentTaskCount >= Task_Execution_Quantity)
                    {
                        return;
                    }
                    currentTaskCount++;
                }

                int totalSeconds = Task_Execution_Duration;
                int numberOfIterations = totalSeconds / 2;
                double increment = 100.0 / numberOfIterations;

                int taskId = GetTaskIdFromTaskName(taskName);

                double currentPercentage = 0;
                for (int i = 1; i <= numberOfIterations; i++)
                {
                    currentPercentage += increment;

                    //using (var tasksKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(tasksRegistryPath, true))
                    //{
                    //    tasksKey.SetValue($"Task_{taskId.ToString("D4")}-Queued", (int)currentPercentage, RegistryValueKind.DWord);
                    //}

                    //string newTaskName = $"Task_{taskId.ToString("D4")}-[" + new string('8', (int)(currentPercentage / 12.5)) + new string('0', 8 - (int)(currentPercentage / 12.5)) + "]-In progress - " + (int)currentPercentage + " percents";
                    //key.SetValue(newTaskName, key.GetValue(taskName), RegistryValueKind.QWord);
                    string progress = new string('8', (int)(currentPercentage / 6.25)) + new string('0', 16 - (int)(currentPercentage / 6.25));
                    char[] chars = progress.ToCharArray();
                    Array.Reverse(chars);
                    string reversed = new string(chars);
                    string newTaskName = $"Task_{taskId.ToString("D4")}-In progress";
                    key.SetValue(newTaskName, reversed, RegistryValueKind.String);
                    //key.DeleteValue(taskName);
                    key.DeleteValue($"Task_{taskId.ToString("D4")}-Queued", false);
                    taskName = newTaskName;

                    Thread.Sleep(2000);
                }

                //string completedTaskName = $"Task_{taskId.ToString("D4")}-[{new string('8', 8)}]-Completed";

                key.DeleteValue(taskName);

                using (var tasksKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(tasksRegistryPath, true))
                {
                    //tasksKey.DeleteValue($"Task_{taskId.ToString("D4")}-Queued", false); // Видалення значення "Queued"
                    tasksKey.SetValue($"Task_{taskId.ToString("D4")}-Completed", "8888888888888888", RegistryValueKind.String); // Створення нового значення "Completed"
                    WriteLog($"Задача Task_{taskId.ToString("D4")} успішно ВИКОНАНА!");
                }
                lock (lockObj)
                {
                    currentTaskCount--;
                }

            });
            key.Close();
        }

        private static int GetTaskIdFromTaskName(string taskName)
        {
            var match = Regex.Match(taskName, @"Task_(\d{4})");
            if (match.Success)
            {
                return int.Parse(match.Groups[1].Value);
            }
            return -1;
        }

        private static void WriteLog(string e)
        {
            using (StreamWriter F = new StreamWriter(logPath, true))
            {
                F.WriteLine("=============================== " + DateTime.Now + " ===============================\n" + e + "\n");
            }
        }

        private string ReadRegistry(string regPath, string parameter)
        {
            using (RegistryKey key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(regPath))
            {
                if (key != null)
                {
                    object val = key.GetValue(parameter);
                    return val?.ToString();
                }
            }
            return null;
        }
    }
}
