using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Filesystem.Ntfs;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace MonElite
{
    public class Program
    {
        private static void Main(string[] args)
        {
            var sendingModel = new ReportModel
            {
                MachineId = Environment.MachineName
            };

            var mode = args.Length == 1 ? args[0].ToLower() : "all";
            Console.WriteLine($"Creating {mode} report...");
            switch (mode)
            {
                case "d":
                case "space-disk":
                    sendingModel.DisksSpaceBytes = GetUsedSpacePerDrive();
                    break;

                case "u":
                case "space-user":
                    sendingModel.UsersSpaceBytes = GetUserProfileFoldersSizes();
                    break;

                case "i":
                case "installed-apps":
                    sendingModel.InstalledApps = GetInstalledApps();
                    break;

                case "l":
                case "users-list":
                    sendingModel.Users = GetUsers();
                    break;

                case "all":
                    sendingModel.Users = GetUsers();
                    sendingModel.InstalledApps = GetInstalledApps();
                    sendingModel.DisksSpaceBytes = GetUsedSpacePerDrive();
                    sendingModel.UsersSpaceBytes = GetUserProfileFoldersSizes();

                    break;

                default:
                    PrintHelp();
                    break;
            }

            Console.WriteLine($"Sending {mode} report...");
            SendReport(sendingModel);
            
            Console.WriteLine("Sent.");
        }

        private static List<string> GetUsers()
        {
            var output = StartProcessAndReadOutput("net", "users");

            var usernames = Regex.Match(output, @"\-\r\n(.*?)The command completed successfully.",
                    RegexOptions.Singleline)
                .Groups[1].Value
                .Split(new[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .ToList();

            return usernames;
        }

        private static void SendReport(ReportModel sendingModel)
        {
            var compiledReport = JsonConvert.SerializeObject(sendingModel);

            var httpClient = new HttpClient();
            var t = httpClient.PostAsync(
                "https://script.google.com/macros/s/AKfycbwIl7ZDB3v8RxYjc4gzna2lwuWhf9YEb3OqeJRG8w0AqdYcbg/exec",
                new StringContent(compiledReport)).Result.Content.ReadAsStringAsync().Result;
        }

        private static List<string> GetInstalledApps()
        {
            var specialMessageWords = new[]
            {
                "Validation Warnings", "A pending system", "being ignored due to the current", "It is recom mended",
                "packages installed.", "not managed with Chocolatey.", "success", "It is recommended that you",
                "Did you know Pro", "Learn more about", "https://chocolatey.org/compare"
            };
            
            var output = StartProcessAndReadOutput("choco", "list -li");

            var installedApps = output
                .Split("\r\n")
                .Where(line => line.Length > 1 && !specialMessageWords.Any(line.Contains))
                .ToList();

            return installedApps;
        }

        private static Dictionary<string, long[]> GetUsedSpacePerDrive()
        {
            var result = DriveInfo.GetDrives()
                .Where(x => x.IsReady)
                .ToDictionary(x => x.Name, x => new[] { x.AvailableFreeSpace, x.TotalSize });

            return result;
        }

        private static Dictionary<string, ulong> GetUserProfileFoldersSizes(string usersPath = "C:\\Users")
        {
            var driveToAnalyze = new DriveInfo(Path.GetPathRoot(usersPath));
            var trimmedUserPath = usersPath.Replace('/', '\\').Trim('"', '\\');

            Dictionary<string, ulong> nodes;
            using (var ntfsReader = new NtfsReader(driveToAnalyze, RetrieveMode.Minimal))
            {
                nodes = ntfsReader.GetNodes(usersPath)
                    .Skip(1)
                    .AsParallel()
                    .Where(node => (node.Attributes & Attributes.Directory) == 0)
                    .ToLookup(x => x.FullName
                        .Substring(trimmedUserPath.Length)
                        .Split('\\', StringSplitOptions.RemoveEmptyEntries)
                        .FirstOrDefault() ?? ".")
                    .ToDictionary(
                        x => x.Key,
                        x => x.Aggregate<INode, ulong>(0, (a, c) => a + c.Size));
            }

            return nodes;
        }

        private static void PrintHelp()
        {
            Console.WriteLine("Sends computer's health status to preconfigurated url");
            Console.WriteLine("help - this text");
            Console.WriteLine("space-disk, d - fill and send 'space per disk' report only");
            Console.WriteLine("space-user, u  - fill and send 'space spelled per user' report only");
            Console.WriteLine("user-list, l  - fill and send 'users' report only");
            Console.WriteLine("installed-apps, i - fill and send 'installed apps' report only");
            Console.WriteLine("all (or empty) - fill and send all reports");
        }


        private static string StartProcessAndReadOutput(string path, string args)
        {
            var p = new Process
            {
                StartInfo =
                {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    FileName = path,
                    Arguments = args
                }
            };
            p.Start();
            var output = p.StandardOutput.ReadToEnd();
            p.WaitForExit();

            return output;
        }
    }
}