class ReportModel {
    ComputerName: string;

    Apps: string[];
    Users: string[];
    Policies: GroupPolicyModel[];
    DriveUsage: DriveUsageModel[];
    ProfilesSizes: ProfilesSizesModel[];
}

class ProfilesSizesModel {
    Username: string;
    UsedSpaceBytes: number;
}

class GroupPolicyModel {
    Area: string;
    Name: string;
    Path: string;
    Value: string;
}

class DriveUsageModel {
    Drive: string;
    FreeSpaceBytes: number;
    TotalSpaceBytes: number;
};