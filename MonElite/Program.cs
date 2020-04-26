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
            ReportModel sendingModel = null;
            ReportBuilder reportBuilder = new ReportBuilder();

            var mode = args.Length == 1 ? args[0].ToLower() : "all";
            Console.WriteLine($"Creating {mode} report...");
            switch (mode)
            {
                case "d":
                case "space-disk":
                    sendingModel = reportBuilder.BuildReport(ReportType.DrivesUsage);
                    break;

                case "u":
                case "space-user":
                    sendingModel = reportBuilder.BuildReport(ReportType.ProfilesSizes);
                    break;

                case "i":
                case "installed-apps":
                    sendingModel = reportBuilder.BuildReport(ReportType.Apps);
                    break;

                case "l":
                case "users-list":
                    sendingModel = reportBuilder.BuildReport(ReportType.Users);
                    break;

                case "p":
                case "policies-list":
                    sendingModel = reportBuilder.BuildReport(ReportType.Policies);
                    break;

                case "all":
                    sendingModel = reportBuilder.BuildReport();
                    break;

                default:
                    PrintHelp();
                    break;
            }

            Console.WriteLine($"Sending {mode} report...");
            SendReport(sendingModel);

            Console.WriteLine("Sent.");
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

        private static void PrintHelp()
        {
            Console.WriteLine("Sends computer's health status to preconfigurated url");
            Console.WriteLine("help - this text");
            Console.WriteLine("space-disk, d - fill and send 'space per disk' report only");
            Console.WriteLine("space-user, u  - fill and send 'space spelled per user' report only");
            Console.WriteLine("user-list, l  - fill and send 'users' report only");
            Console.WriteLine("installed-apps, i - fill and send 'installed apps' report only");
            Console.WriteLine("policies-list, p - fill and send 'local group policies' report only");
            Console.WriteLine("all (or empty) - fill and send all reports");
        }
    }
}