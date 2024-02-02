#-----------------------------------------------------------------------
# WSUSOfflineScan 
# Frank Maxwitat, 
# Last Update Feb 01, 2023
#-----------------------------------------------------------------------

$Version = 'Version 1.0.3'
$CabPath = "$PSScriptRoot\wsusscn2.cab"
$logFile = "$env:windir\Logs\WSUSOfflineCatalog.log"
#Important: Read the digital signature date from the WSUSScn2.cab and update this date below. 
# To get the signature timestamp: Right-click the cab, open the Digital Signature tab and click on detail
$WSUSScn2SignatureDate = "09.01.2024"

#------------------ Begin Functions ------------------------------------
function Log([string]$ContentLog) 
{
    Write-Host "$(Get-Date -Format "dd.MM.yyyy HH:mm:ss,ff") $($ContentLog)"
    Add-Content -Path $logFile -Value "$(Get-Date -Format "dd.MM.yyyy HH:mm:ss,ff") $($ContentLog)"
}

# Creates a new class in WMI to store our data
function CreateMissingWSUSUpdatesWMIClass()
{
    $newClass = New-Object System.Management.ManagementClass("root\cimv2", [String]::Empty, $null);

    $newClass["__CLASS"] = "MissingWSUSUpdates";

    $newClass.Qualifiers.Add("Static", $true)
    $newClass.Properties.Add("kbNumber", [System.Management.CimType]::String, $false)
    $newClass.Properties["kbNumber"].Qualifiers.Add("key", $true)
    $newClass.Properties["kbNumber"].Qualifiers.Add("read", $true)
    $newClass.Properties.Add("Title", [System.Management.CimType]::String, $false)
    $newClass.Properties["Title"].Qualifiers.Add("read", $true)
    $newClass.Properties.Add("MsrcSeverity", [System.Management.CimType]::String, $false)
    $newClass.Properties["MsrcSeverity"].Qualifiers.Add("read", $true)
    $newClass.Properties.Add("Categories", [System.Management.CimType]::String, $false)
    $newClass.Properties["Categories"].Qualifiers.Add("read", $true)
    $newClass.Properties.Add("LastChangeTime", [System.Management.CimType]::String, $false)
    $newClass.Properties["LastChangeTime"].Qualifiers.Add("read", $true)
    $newClass.Properties.Add("CabSignatureTimestamp", [System.Management.CimType]::String, $false)
    $newClass.Properties["CabSignatureTimestamp"].Qualifiers.Add("read", $true)
    $newClass.Put()
}

#------------------ End Functions --------------------------------------

if(Test-Path $logFile){Remove-Item $logfile}

Log "Stating Script $Version"

# Check if cab exists
if(!(Test-Path $CabPath))
{
    Log "Error: Can't find wsusscn2.cab at $CabPath"
    exit 1
}
else{
    $cabLastWriteTime = ((Get-Item -Path $CabPath).LastWriteTime)
}

Log "Creating Windows Update session"
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateServiceManager  = New-Object -ComObject Microsoft.Update.ServiceManager 

$UpdateService = $UpdateServiceManager.AddScanPackageService("Offline Sync Service", $CabPath, 1) 

Log "Creating Windows Update Searcher"
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()  
$UpdateSearcher.ServerSelection = 3
$UpdateSearcher.ServiceID = $UpdateService.ServiceID.ToString()
 
Log "Searching for missing updates"
$SearchResult = $UpdateSearcher.Search("IsInstalled=0")

$Updates = $SearchResult.Updates
#Optional: Output as csv
#$Updates | Export-Csv -Path $env:windir\Logs\MissingWSUSUpdates.csv -Force

Log (([string]($Updates.Count)) + " updates missing")

#------------------
$UpdatesSummary = If ($searchresult.Updates.Count  -gt 0) 
{
        #Updates that are missing
        $count  = $searchresult.Updates.Count
   
        For ($i=0; $i -lt $Count; $i++) 
        {
            $Update  = $searchresult.Updates.Item($i)

            [pscustomobject]@{
            Title =  $Update.Title
            KB =  $($Update.KBArticleIDs)
            Severity = $($Update.MsrcSeverity)            
            Categories = [string](($Update.Categories | Select-Object -ExpandProperty Name )-join '; ')
            LastChangeTime = $Update.LastDeploymentChangeTime.ToString("MM/dd/yyyy")
            CabSignatureTimestamp = $WSUSScn2SignatureDate
        }
    }       
}
else
{
    Log "No missing updates found"
}

Log "Checking whether we already created our custom WMI class MissingWSUSUpdates on this PC, if not, we'll do"
[void](Get-WMIObject MissingWSUSUpdates -ErrorAction SilentlyContinue -ErrorVariable wmiclasserror)
if ($wmiclasserror)
{
    try { CreateMissingWSUSUpdatesWMIClass }
    catch
    {
        Log "Could not create WMI class"
        Exit 1
    }
}
else{
    Log "WMI Class MissingWSUSUpdates created"
}

Log "Clearing WMI"
Get-WmiObject MissingWSUSUpdates | Remove-WmiObject

Log "Storing the missing updates information in WMI"
for ($i=0; $i -lt $UpdatesSummary.Count; $i++)
{
    [void](Set-WmiInstance -Path \\.\root\cimv2:MissingWSUSUpdates -Arguments @{kbNumber=$UpdatesSummary[$i].KB; Title=$UpdatesSummary[$i].Title; `
    MsrcSeverity=$UpdatesSummary[$i].MsrcSeverity; Categories=$UpdatesSummary[$i].Categories; LastChangeTime=$UpdatesSummary[$i].LastChangeTime; CabSignatureTimestamp = $UpdatesSummary[$i].CabSignatureTimestamp})
}
Log "Open wbemtest, connect to root\cimv2 and run 'select * from MissingWSUSUpdates' to check the WMI."

Log "Finishing"
