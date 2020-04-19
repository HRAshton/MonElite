using System.Collections.Generic;

namespace MonElite
{
    internal class ReportModel
    {
        public Dictionary<string, long[]> DisksSpaceBytes { get; set; } = new Dictionary<string, long[]>();

        public Dictionary<string, ulong> UsersSpaceBytes { get; set; } = new Dictionary<string, ulong>();

        public List<string> InstalledApps { get; set; } = new List<string>();

        public string MachineId { get; set; } = "";

        public List<string> Users { get; set; } = new List<string>();
    }
}