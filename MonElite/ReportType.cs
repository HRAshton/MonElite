namespace MonElite
{
    public enum ReportType
    {
        Apps = 1 << 0,
        Users = 1 << 1,
        Policies = 1 << 2,
        DrivesUsage = 1 << 3,
        ProfilesSizes = 1 << 4
    }
}