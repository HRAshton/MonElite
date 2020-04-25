using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Filesystem.Ntfs;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
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
                    sendingModel.DisksSpaceBytes = GetUsedSpacePerDriveAsync().Result;
                    break;

                case "u":
                case "space-user":
                    sendingModel.UsersSpaceBytes = GetUserProfileFoldersSizesAsync().Result;
                    break;

                case "i":
                case "installed-apps":
                    sendingModel.Apps = GetAppsAsync().Result;
                    break;

                case "l":
                case "users-list":
                    sendingModel.Users = GetUsersAsync().Result;
                    break;

                case "all":
                    Task<List<string>> appsTask = GetAppsAsync();
                    Task<List<string>> usersTask = GetUsersAsync();
                    Task<List<string[]>> policiesTask = GetLocalGroupPoliciesAsync();
                    Task<Dictionary<string, long[]>> drivesTask = GetUsedSpacePerDriveAsync();
                    Task<Dictionary<string, ulong>> usersFoldersTask = GetUserProfileFoldersSizesAsync();

                    //sendingModel.Apps = appsTask.Result;
                    //sendingModel.Users = usersTask.Result;
                    sendingModel.Policies = policiesTask.Result;
                    //sendingModel.DisksSpaceBytes = drivesTask.Result;
                    //sendingModel.UsersSpaceBytes = usersFoldersTask.Result;

                    break;

                default:
                    PrintHelp();
                    break;
            }

            Console.WriteLine($"Sending {mode} report...");
            SendReport(sendingModel);

            Console.WriteLine("Sent.");
        }

        private async static Task<List<string>> GetUsersAsync()
        {
            var output = await StartProcessAndReadOutputAsync("net", "users");

            var usernames = Regex.Match(output, @"\-{10}\r{0,1}\n(.*?)\r{0,1}\n([а-яА-Я\w ]+)\.",
                    RegexOptions.Singleline)
                .Groups[1].Value
                .Split(new[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .ToList();

            return usernames;
        }

        private static void SendReport(ReportModel sendingModel)
        {
            var compiledReport = JsonConvert.SerializeObject(sendingModel);

            Console.WriteLine("Report content: " + compiledReport);
            var httpClient = new HttpClient();
            var response = httpClient.PostAsync(
                "https://script.google.com/macros/s/AKfycbwIl7ZDB3v8RxYjc4gz-na2lwuWhf9YEb3OqeJRG8w0AqdYcbg/exec",
                new StringContent(compiledReport))
                .Result.Content.ReadAsStringAsync().Result;
            Console.WriteLine("Response: " + response);
        }

        private async static Task<List<string>> GetAppsAsync()
        {
            var specialMessageWords = new[]
            {
                "Validation Warnings", "A pending system", "being ignored due to the current", "It is recom mended",
                "packages installed.", "not managed with Chocolatey.", "success", "It is recommended that you",
                "Did you know Pro", "Learn more about", "https://chocolatey.org/compare"
            };

            var output = await StartProcessAndReadOutputAsync("choco", "list -li");

            var installedApps = output
                .Split("\r\n")
                .Where(line => line.Length > 1 && !specialMessageWords.Any(line.Contains))
                .OrderBy(x => x)
                .ToList();

            return installedApps;
        }

        private async static Task<Dictionary<string, long[]>> GetUsedSpacePerDriveAsync()
        {
            return await Task.Run(() =>
            {
                var result = DriveInfo.GetDrives()
                    .Where(x => x.IsReady && x.DriveType == DriveType.Fixed)
                    .OrderBy(x => x.Name)
                    .ToDictionary(x => x.Name.Substring(0, 1), x => new[] { x.AvailableFreeSpace, x.TotalSize });

                return result;
            });
        }

        private async static Task<Dictionary<string, ulong>> GetUserProfileFoldersSizesAsync(string usersPath = "C:\\Users")
        {
            return await Task.Run(() =>
            {
                var driveToAnalyze = new DriveInfo(Path.GetPathRoot(usersPath));
                var trimmedUserPath = usersPath.Replace('/', '\\').Trim('"', '\\');

                Dictionary<string, ulong> nodes;
                using (var ntfsReader = new NtfsReader(driveToAnalyze, RetrieveMode.Minimal))
                {
                    nodes = ntfsReader.GetNodes(usersPath)
                        .AsParallel()
                        .Where(node => (node.Attributes & Attributes.Directory) == 0)
                        .ToLookup(x => x.FullName
                            .Substring(trimmedUserPath.Length)
                            .Split('\\')
                            .ElementAtOrDefault(1) ?? ".")
                        .Where(x => !x.Key.Contains(".LOG", StringComparison.OrdinalIgnoreCase))
                        .OrderBy(x => x.Key)
                        .ToDictionary(
                            x => x.Key,
                            x => x.Aggregate<INode, ulong>(0, (a, c) => a + c.Size));
                }

                return nodes;
            });
        }

        private async static Task<List<string[]>> GetLocalGroupPoliciesAsync()
        {
            var tempFolder = Directory.CreateDirectory(Path.GetTempPath() + Guid.NewGuid());

            var lgpoOutput = await StartProcessAndReadOutputAsync("lgpo.exe", "/b " + tempFolder.FullName);
            var polFiles = tempFolder.GetFiles("*.pol", SearchOption.AllDirectories);
            var outputs = polFiles
                .Select(file => StartProcessAndReadOutputAsync("lgpo.exe", "/parse /m " + file.FullName).Result)
                .ToList();
            var blocks = outputs.SelectMany(output => output.Split(Environment.NewLine + Environment.NewLine));
            var result = blocks
                .Select(block => block.Split(Environment.NewLine))
                .Where(lines => lines.Length == 4)
                .Select(lines => new[] {
                    lines[2],
                    lines[3].Split(":").Last(),
                    lines[1]
                })
                .OrderBy(x => x[2]).ThenBy(x => x[0])
                .ToList();

            tempFolder.Delete(true);

            return result;
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


        private async static Task<string> StartProcessAndReadOutputAsync(string path, string args)
        {
            return await Task.Run(() =>
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
            });
        }
    }
}