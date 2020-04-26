function doPost(e: GoogleAppsScript.Events.DoPost): void {
    const reportData = JSON.parse(e.postData.contents) as ReportModel;
    // var reportData = JSON.parse('{"DisksSpaceBytes":{},"UsersSpaceBytes":{},"Apps":[],"MachineId":"W03307","Users":[],"Policies":[["NoClose","1","Software/Microsoft/Windows/CurrentVersion/Policies/Explorer"],["NoClose","1","Software/Microsoft/Windows/CurrentVersion/Policies/Explorer"],["*","DELETEALLVALUES","Software/Policies/Google/Chrome/ExtensionInstallForcelist"],["1","//clients2.google.com/service/update2/crx","Software/Policies/Google/Chrome/ExtensionInstallForcelist"],["Category","1","Software/Policies/Microsoft/Windows NT/CurrentVersion/NetworkList/Signatures/010103000F0000F0080000000F0000F08E8F727A88F84DDEFE9C1C2BB78CC9A1AD049579F7FE4CF02C5E5F83A47CBDB3"],["CategoryReadOnly","1","Software/Policies/Microsoft/Windows NT/CurrentVersion/NetworkList/Signatures/010103000F0000F0080000000F0000F08E8F727A88F84DDEFE9C1C2BB78CC9A1AD049579F7FE4CF02C5E5F83A47CBDB3"],["Icon16","","Software/Policies/Microsoft/Windows NT/CurrentVersion/NetworkList/Signatures/010103000F0000F0080000000F0000F08E8F727A88F84DDEFE9C1C2BB78CC9A1AD049579F7FE4CF02C5E5F83A47CBDB3"],["Icon24","","Software/Policies/Microsoft/Windows NT/CurrentVersion/NetworkList/Signatures/010103000F0000F0080000000F0000F08E8F727A88F84DDEFE9C1C2BB78CC9A1AD049579F7FE4CF02C5E5F83A47CBDB3"],["Icon32","","Software/Policies/Microsoft/Windows NT/CurrentVersion/NetworkList/Signatures/010103000F0000F0080000000F0000F08E8F727A88F84DDEFE9C1C2BB78CC9A1AD049579F7FE4CF02C5E5F83A47CBDB3"],["Icon48","","Software/Policies/Microsoft/Windows NT/CurrentVersion/NetworkList/Signatures/010103000F0000F0080000000F0000F08E8F727A88F84DDEFE9C1C2BB78CC9A1AD049579F7FE4CF02C5E5F83A47CBDB3"],["IconReadOnly","1","Software/Policies/Microsoft/Windows NT/CurrentVersion/NetworkList/Signatures/010103000F0000F0080000000F0000F08E8F727A88F84DDEFE9C1C2BB78CC9A1AD049579F7FE4CF02C5E5F83A47CBDB3"],["NameReadOnly","1","Software/Policies/Microsoft/Windows NT/CurrentVersion/NetworkList/Signatures/010103000F0000F0080000000F0000F08E8F727A88F84DDEFE9C1C2BB78CC9A1AD049579F7FE4CF02C5E5F83A47CBDB3"],["NetworkName","VipNET Client (SuntsovK Tomsk)","Software/Policies/Microsoft/Windows NT/CurrentVersion/NetworkList/Signatures/010103000F0000F0080000000F0000F08E8F727A88F84DDEFE9C1C2BB78CC9A1AD049579F7FE4CF02C5E5F83A47CBDB3"],["uNCName","////w03726//","Software/Policies/Microsoft/Windows NT/Printers/PushedConnections/{21A01897-D7C9-40DE-ACB9-5E95132603C7}"],["HidePeopleBar","1","Software/Policies/Microsoft/Windows/Explorer"],["HidePeopleBar","1","Software/Policies/Microsoft/Windows/Explorer"],["EnableScripts","1","Software/Policies/Microsoft/Windows/PowerShell"],["ExecutionPolicy","AllSigned","Software/Policies/Microsoft/Windows/PowerShell"],["UseOEMBackground","1","Software/Policies/Microsoft/Windows/System"],["ActiveHoursEnd","23","Software/Policies/Microsoft/Windows/WindowsUpdate"],["ActiveHoursStart","8","Software/Policies/Microsoft/Windows/WindowsUpdate"],["SetActiveHours","1","Software/Policies/Microsoft/Windows/WindowsUpdate"],["AllowMUUpdateService","DELETE","Software/Policies/Microsoft/Windows/WindowsUpdate/AU"],["AUOptions","4","Software/Policies/Microsoft/Windows/WindowsUpdate/AU"],["AutomaticMaintenanceEnabled","DELETE","Software/Policies/Microsoft/Windows/WindowsUpdate/AU"],["DetectionFrequency","6","Software/Policies/Microsoft/Windows/WindowsUpdate/AU"],["DetectionFrequencyEnabled","1","Software/Policies/Microsoft/Windows/WindowsUpdate/AU"],["NoAutoUpdate","0","Software/Policies/Microsoft/Windows/WindowsUpdate/AU"],["ScheduledInstallDay","0","Software/Policies/Microsoft/Windows/WindowsUpdate/AU"],["ScheduledInstallEveryWeek","1","Software/Policies/Microsoft/Windows/WindowsUpdate/AU"],["ScheduledInstallFirstWeek","DELETE","Software/Policies/Microsoft/Windows/WindowsUpdate/AU"],["ScheduledInstallFourthWeek","DELETE","Software/Policies/Microsoft/Windows/WindowsUpdate/AU"],["ScheduledInstallSecondWeek","DELETE","Software/Policies/Microsoft/Windows/WindowsUpdate/AU"],["ScheduledInstallThirdWeek","DELETE","Software/Policies/Microsoft/Windows/WindowsUpdate/AU"],["ScheduledInstallTime","22","Software/Policies/Microsoft/Windows/WindowsUpdate/AU"]]}');

    handle(reportData);
}

function handle(reportData: ReportModel): void {
    const compName = reportData.ComputerName.toLowerCase();

    if (reportData.Apps && reportData.Apps.length)
        fillInstalledApps(compName, reportData.Apps);

    if (reportData.Users && reportData.Users.length)
        fillUsers(compName, reportData.Users);

    if (reportData.Policies && reportData.Policies.length)
        fillPolicies(compName, reportData.Policies);

    if (reportData.DriveUsage && reportData.DriveUsage.length)
        fillSpacePerDrive(compName, reportData.DriveUsage);

    if (reportData.ProfilesSizes && reportData.ProfilesSizes.length)
        fillUserProfileFoldersSizes(compName, reportData.ProfilesSizes);
}



function fillInstalledApps(compName: string, appsList: string[]): void {
    FillColumnsWithData(compName, ['A', 'A'], appsList.map(a => [a]));
}

function fillUsers(compName: string, usersList: string[]): void {
    FillColumnsWithData(compName, ['B', 'B'], usersList.map(u => [u]));
}

function fillPolicies(compName: string, policiesList: GroupPolicyModel[]) {
    const formattedData = policiesList
        .map(p => [`${p.Area}:${p.Path}/${p.Name} = ${p.Value}`]);

    FillColumnsWithData(compName, ['H', 'H'], formattedData);
}

function fillSpacePerDrive(compName: string, drivesDict: DriveUsageModel[]) {
    const formattedData = drivesDict
        .map(d => [d.Drive, d.FreeSpaceBytes, d.TotalSpaceBytes]);

    FillColumnsWithData(compName, ['C', 'E'], formattedData);
}

function fillUserProfileFoldersSizes(compName: string, userSpaceDict: ProfilesSizesModel[]) {
    const formattedData = userSpaceDict
        .map(d => [d.Username, d.UsedSpaceBytes])
        .sort((a, b) => (a[0] as string).localeCompare(b[0] as string));

    FillColumnsWithData(compName, ['F', 'G'], formattedData);
}


function FillColumnsWithData(compName: string, colsNames: string[], data: any[][]) {
    const headerRows = 1;

    const inftyRange = `${compName}!${colsNames[0]}${headerRows + 1}:${colsNames[1]}`;
    SpreadsheetApp.getActiveSpreadsheet()
        .getRange(inftyRange)
        .clear();

    const sortedData = colsNames[0] === 'A' ? data : data.sort((a, b) => a[0].localeCompare(b[0]));
    SpreadsheetApp.getActiveSpreadsheet()
        .getRange(`${inftyRange}${data.length + headerRows}`)
        .setValues(sortedData);
}