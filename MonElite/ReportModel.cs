using System.Collections.Generic;

namespace MonElite
{
    class ReportModel
    {
        public string ComputerName { get; set; } = "";

        public List<string> Apps { get; set; } = new List<string>();
        public List<string> Users { get; set; } = new List<string>();
        public List<GroupPolicyModel> Policies { get; set; } = new List<GroupPolicyModel>();
        public List<DriveSpaceModel> DrivesUsage { get; set; } = new List<DriveSpaceModel>();
        public List<UserSpaceModel> ProfilesSizes { get; set; } = new List<UserSpaceModel>();
    }

    class UserSpaceModel
    {
        public string Username;
        public ulong UsedSpaceBytes;
    }

    class GroupPolicyModel
    {
        public string Area;
        public string Name;
        public string Path;
        public string Value;
    }

    class DriveSpaceModel
    {
        public string Drive;
        public ulong FreeSpaceBytes;
        public ulong TotalSpaceBytes;
    };
}