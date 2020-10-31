using System;
using System.Management;
using System.Diagnostics;
using System.IO;
using WUApiLib;
using System.Reflection;
using Microsoft.Win32;
using System.Security.Principal;

namespace ListSystemUpdates
{

    class Program
    {

        private static void AdminRelauncher()
        {
            if (!IsRunAsAdmin())
            {
                ProcessStartInfo proc = new ProcessStartInfo();
                proc.UseShellExecute = true;
                proc.WorkingDirectory = Environment.CurrentDirectory;
                proc.FileName = Assembly.GetEntryAssembly().CodeBase;
                proc.Verb = "runas";
                try
                {
                    Process.Start(proc);
                    Process.GetCurrentProcess().Kill();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("This program must be run as an administrator! \n\n" + ex.ToString());
                }
            }
        }

        private static bool IsRunAsAdmin()
        {
            WindowsIdentity id = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(id);

            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private static void InstalledUpdatesFallback()
        {
            const string query = "SELECT HotFixID FROM Win32_QuickFixEngineering";
            var search = new ManagementObjectSearcher(query);
            var collection = search.Get();

            foreach (ManagementObject quickFix in collection)
                Console.WriteLine(quickFix["HotFixID"].ToString());
        }

        private static string GetMachineName()
        {
            return Environment.MachineName;
    }

        private static void InstalledUpdates()
        {
            UpdateSession UpdateSession = new UpdateSession();
            IUpdateSearcher UpdateSearchResult = UpdateSession.CreateUpdateSearcher();
            UpdateSearchResult.Online = false;
            ISearchResult SearchResults = UpdateSearchResult.Search("IsInstalled=1 AND IsHidden=0");
            //http://msdn.microsoft.com/en-us/library/windows/desktop/aa386526(v=VS.85).aspx
             foreach (IUpdate x in SearchResults.Updates)
            {
                Console.WriteLine($"{x.Title} issued {x.LastDeploymentChangeTime}");
            }
        }
        private static string GetOSVersionString()
        {
            var os = Environment.OSVersion;
            return os.VersionString;
        }

        private static string GetReleaseId()
        {
            return Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ReleaseId", "").ToString();
        }

        private static string GetEditionId()
        {
            return Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "EditionId", "").ToString();
        }

        static void Main(string[] args)
        {
            AdminRelauncher();
            Console.WriteLine("Collecting, please wait. Result will be saved in Result.txt");
            Console.WriteLine("Collecting installed Hot fixes");
            FileStream fs = new FileStream("Result.txt", FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            Console.SetOut(sw);
            Console.WriteLine("--------Windows info--------");
            Console.WriteLine($"Machine name: {GetMachineName()}");
            Console.Write($"{GetOSVersionString()} {GetEditionId()}");
            var releaseId = GetReleaseId();
            if (!string.IsNullOrEmpty(releaseId))
                Console.WriteLine($"Release Id: {releaseId}");
            else Console.WriteLine();
                Console.WriteLine("--------Windows updates list--------");
            try
            {
                InstalledUpdates();
            }      
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to fetch updates via WUApi because of {ex.Message}");
                Console.WriteLine("Getting KB list using WMI");
                InstalledUpdatesFallback();
            }    
            sw.Close();
        }
    }
}
