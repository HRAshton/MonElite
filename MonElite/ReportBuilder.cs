using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Filesystem.Ntfs;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MonElite
{
    class ReportBuilder
    {
        public ReportModel BuildReport(params ReportType[] reportTypes)
        {
            if (!reportTypes.Any())
            {
                reportTypes = Enum.GetValues(typeof(ReportType)).Cast<ReportType>().ToArray();
            }

            ReportModel reportModel = new ReportModel { ComputerName = Environment.MachineName };
            List<Task> tasks = new List<Task>();
            if (reportTypes.Contains(ReportType.Apps))
            {
                tasks.Add(GetAppsAsync().ContinueWith(res => reportModel.Apps = res.Result));
            }
            if (reportTypes.Contains(ReportType.Users))
            {
                tasks.Add(GetUsersAsync().ContinueWith(res => reportModel.Users = res.Result));
            }
            if (reportTypes.Contains(ReportType.Policies))
            {
                tasks.Add(GetLocalGroupPoliciesAsync().ContinueWith(res => reportModel.Policies = res.Result));
            }
            if (reportTypes.Contains(ReportType.DrivesUsage))
            {
                tasks.Add(GetDrivesUsageAsync().ContinueWith(res => reportModel.DrivesUsage = res.Result));
            }
            if (reportTypes.Contains(ReportType.ProfilesSizes))
            {
                tasks.Add(GetProfilesSizesAsync().ContinueWith(res => reportModel.ProfilesSizes = res.Result));
            }

            Task.WaitAll(tasks.ToArray());

            return reportModel;
        }

        private async Task<List<string>> GetAppsAsync()
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
                .ToList();

            return installedApps;
        }

        private async Task<List<string>> GetUsersAsync()
        {
            var output = await StartProcessAndReadOutputAsync("net", "users");

            var usernames = Regex.Match(output, @"\-{10}\r{0,1}\n(.*?)\r{0,1}\n([а-яА-Я\w ]+)\.",
                    RegexOptions.Singleline)
                .Groups[1].Value
                .Split(new[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .ToList();

            return usernames;
        }

        private async Task<List<GroupPolicyModel>> GetLocalGroupPoliciesAsync()
        {
            var tempFolder = Directory.CreateDirectory(Path.GetTempPath() + Guid.NewGuid());

            var lgpoOutput = await StartProcessAndReadOutputAsync("lgpo.exe", "/b " + tempFolder.FullName);
            var polFiles = tempFolder.GetFiles("*.pol", SearchOption.AllDirectories);

            var policies = polFiles
                .SelectMany(file =>
                {
                    var parsedPolRawText = StartProcessAndReadOutputAsync("lgpo.exe", "/parse /m " + file.FullName).Result;
                    var objectBlocks = parsedPolRawText.Split(Environment.NewLine + Environment.NewLine);
                    var policies = objectBlocks
                        .Select(block => block.Split(Environment.NewLine))
                        .Where(lines => lines.Length == 4)
                        .Select(lines => new GroupPolicyModel
                        {
                            Area = file.Directory.Name.Substring(0, 2),
                            Path = lines[1].Replace('\\', '/'),
                            Name = lines[2],
                            Value = lines[3]
                        })
                        .ToList();

                    return policies;
                })
                .ToList();

            tempFolder.Delete(true);

            return policies;
        }

        private async Task<List<DriveSpaceModel>> GetDrivesUsageAsync()
        {
            return await Task.Run(() =>
            {
                var result = DriveInfo.GetDrives()
                    .Where(x => x.IsReady && x.DriveType == DriveType.Fixed)
                    .OrderBy(x => x.Name)
                    .Select(drInfo => new DriveSpaceModel
                    {
                        Drive = drInfo.Name.Substring(0, 1),
                        FreeSpaceBytes = (ulong)drInfo.AvailableFreeSpace,
                        TotalSpaceBytes = (ulong)drInfo.TotalSize
                    })
                    .ToList();

                return result;
            });
        }

        private async Task<List<UserSpaceModel>> GetProfilesSizesAsync(string usersPath = "C:\\Users")
        {
            return await Task.Run(() =>
            {
                var driveToAnalyze = new DriveInfo(Path.GetPathRoot(usersPath));
                var trimmedUserPath = usersPath.Replace('/', '\\').Trim('"', '\\');

                List<UserSpaceModel> nodes;
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
                        .Select(gr => new UserSpaceModel
                        {
                            Username = gr.Key,
                            UsedSpaceBytes = gr.Aggregate<INode, ulong>(0, (a, c) => a + c.Size)
                        })
                        .ToList();
                }

                return nodes;
            });
        }


        private async Task<string> StartProcessAndReadOutputAsync(string path, string args)
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
