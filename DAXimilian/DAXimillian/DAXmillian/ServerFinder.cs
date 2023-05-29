using System.Diagnostics;
using System.IO;
using System.Management;
using System.Text.RegularExpressions;


public class ServerFinder
{
    public static string FindServer()
    {
        string GetPowerBILocalhost()
        {
            var processes = Process.GetProcessesByName("msmdsrv");

            foreach (var process in processes)
            {
                var commandLine = GetCommandLine(process);
                var match = Regex.Match(commandLine ?? "", @"-s ""([^""]+)""", RegexOptions.IgnoreCase);

                if (match.Success)
                {
                    var dataDirectory = match.Groups[1].Value;
                    var portFile = Path.Combine(dataDirectory, "msmdsrv.port.txt");

                    if (File.Exists(portFile))
                    {
                        var portNumber = File.ReadAllText(portFile).Trim();
                        return $"localhost:{portNumber}";
                    }
                }
            }

            return null;
        }

        string GetCommandLine(Process process)
        {
            string commandLine = null;
            using (var searcher = new ManagementObjectSearcher($"SELECT CommandLine FROM Win32_Process WHERE ProcessId = {process.Id}"))
            {
                foreach (var @object in searcher.Get())
                {
                    commandLine += @object["CommandLine"] + " ";
                }
            }

            return commandLine;
        }

        string localhost = GetPowerBILocalhost();

        if (!string.IsNullOrEmpty(localhost))
        {
     
            return Regex.Replace(localhost, @"[^ -~]", ""); ;
        }
        else
        {
            
            return null;
        }
    }

}
