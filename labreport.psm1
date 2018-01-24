param(
    [parameter(Position=0,Mandatory=$false)][boolean]$BeQuiet=$false
)

if (([Environment]::UserInteractive) -and (!$BeQuiet)) {
    write-host "***********************************************************************************"
    write-host " Type 'help COMMAND' for detailed info on each command."
    write-host " Commands:"
    write-host " "
    write-host " Summary            Summary of all labs"
    write-host " Report LAB         Lab report for labs matching LAB"
    write-host " Update-Labs        Force update list of labs from SCCM (automatic every 14 days)"
    write-host " Update-LabReport   Force update lab info (automatic every day when module reloaded"
    write-host " Update             Update list of labs and lab info if required"
    write-host " Publish-Labreport  Make HTML report"
    write-host "***********************************************************************************"
}

$SiteServer = "wsp-configmgr01"

$cachedir="$($env:localappdata)\labreport"
$cachedir="\\generalvs\general\ITServices\Department\Records\Accreditation\current-semester-prep\labreport"

if (-not (Test-Path $cachedir)) { New-Item -Path $cachedir -ItemType directory }
$labscache = "$cachedir\labs.xml"
$LabReportCache = "$cachedir\report.xml"

#$reportdir="c:\scratch\drees\git\dardie\usc-labreport\html"
$reportdir="$cachedir\html"

if (-not (Test-Path $reportdir)) {
    New-Item -Path $reportdir -ItemType directory
    #write-error "Report dir ""$reportdir"" does not exist, please fix."
}

class PCInfo {
    [String]$Collection
    [String]$ComputerName
    [String]$OU
    [DateTime]$InstallDate
    [String]$OSName
    [String]$OSBuild
    [String]$OSVersion
    [DateTime]$LastBoot
    [DateTime]$LastLogon
    [DateTime]$LastPolicyRequest
    [DateTime]$LastActiveDate
    [String]$MacAddresses
    [String]$Manufacturer
    [String]$Model
    [Object]$User
    [Uint32]$Memory
    [UInt32]$VirtualMemory
    [Object[]]$IPSubnets
    [Object[]]$IPAddresses
    [Object[]]$ADGroups
}

class LabSummary {
    [String]$Collection
    [Uint32]$Down7Days
    [UInt32]$TotalPCs
    [UInt32]$Win7
}
function show-module ($modname) {
	<#
     .SYNOPSIS
	 Show a list of all Commands in the specified module.
	 
     .PARAMETER Module
	 Name of module to list commands.
	 
	 .EXAMPLE
	 Show-Module
	 Show-SDCommands USC-SemesterPrep
	 
	 .NOTES
      Author: Darryl Rees
      Date Created: 5 January 2017    
      ChangeLog:
	#>

	if (-not $modname) {
		$modname="(none specified)"
	}
	if (-not (test-path $modname)) {
		$modname2='\\usc\usc\appdev\General\SCCMTools\Scripts\Modules\' + $modname + '\' + $modname + '.psm1'
		if (test-path $modname2) {
			$modname=$modname2
		} else {
			write-output ("No such module " + $modname)
			write-output ("Perhaps you were wanting one of the USC modules?")
			dir "\\usc\usc\appdev\General\SCCMTools\Scripts\Modules\"
			return
		}
	}
	get-module $modname -listavailable | % { ($_.exportedcommands.values.Name) } | % { get-help $_ | select name, synopsis }
}
Export-ModuleMember -function Show-Module
Set-Alias labreports 'show-module lab-reports'
Export-ModuleMember -alias labreports

Function Test-FileOlder ($Path, $Days) {
    if (Test-Path $Path) {
        $lastWrite = (get-item $Path).LastWriteTime 
        $timespan = new-timespan -days $Days
        Return (((get-date) - $lastWrite) -gt $timespan)
    } else {
        Return $False
    }
}

Function Update-CachedDataIfNecessary () {
    # "Update list of labs from SCCM 1/14 days"
    If (-not (test-connection $SiteServer -count 1 -erroraction silentlycontinue)) {
        Write-Warning "Unable to contact SCCM Server $SiteServer, using cached data without updating"
        Return
    }

    $LabsCacheRefreshPeriod = 14
    if (-not (Test-Path $labscache)) {
        Write-Warning "Pulling list of Managed Labs collections from SCCM, please be patient"
        $Labs = Update-Labs
    } elseif (Test-FileOlder $labscache $LabsCacheRefreshPeriod) {
        $Msg = "Managed Labs collections have not been updated from SCCM in over $LabsCacheRefreshPeriod days. Update? [y/n]"
        $Confirm = Read-Host $Msg
        $ForceUpdateLabsCache = ($Confirm -eq "y")
        Update-Labs
    }

    # "Update lab info report from SCCM automatically once / day"
    $LabsReportCacheRefreshPeriod = 1
    if ((-not (Test-Path $LabReportCache)) -or (Test-FileOlder $LabReportCache $LabsReportCacheRefreshPeriod)) {
        Update-LabReport
    }
}
Export-ModuleMember -function Update-CachedDataIfNecessary
Set-Alias Update Update-CachedDataIfNecessary
Export-ModuleMember -alias Update

Function Update-Labs {
    Write-Verbose "Getting list of all Managed Labs from SCCM"
    $Labs = Get-cfgcollectionsByFolder "Managed Labs"
    $Labs | where {$_.collectionname -notlike "_*" -and $_.collectionname -notlike "All Managed Labs"} | Export-CliXML $labscache
}
Export-ModuleMember -function Update-Labs

Function Update-LabReport {
    param(
        [parameter(Mandatory=$false)][String]$LabSearchString
    )

    $Labs = Import-CliXML $LabsCache

    if ($LabSearchString) {
        $LabsToUpdate = $Labs.CollectionName | ? { $_ -like "*$($LabSearchString)*"}
        $Report = Import-CliXML $LabReportCache | where { $LabsToUpdate -notcontains $_.Collection }
    } else {
        $Report = @()
        $LabsToUpdate = $Labs.CollectionName
    }

    $Progress = [ProgressBar]::New("Gathering lab PC info", $LabsToUpdate.count)
    $ReportUpdate = ForEach ($LabName in $LabsToUpdate) {
        $Progress.Advance($LabName)
        Get-LabInfo $LabName
    }
    $Report = $Report + $ReportUpdate
    $Report | Export-CliXML $LabReportCache
}
Export-ModuleMember -function Update-LabReport

Class ProgressBar {
    hidden [int]$CurrentStep = 0
    hidden [int]$Totalsteps = 0
    hidden [String]$Stage = "NO STAGE SPECIFIED IN CONSTRUCTOR"
    Progressbar([String]$Stage, [Int]$TotalSteps) {
        $this.CurrentStep = 0
        $this.TotalSteps = $TotalSteps
        $this.Stage = $Stage
    }
    [Void] Advance ([String]$Status) {
        $this.CurrentStep++
        $PercentComplete = [math]::Round($this.CurrentStep * 100 / $this.TotalSteps)
        if ($PercentComplete -gt 100) { $Percentcomplete = 100 }
        Write-Progress -activity $this.Stage -status ([String]$PercentComplete + "% : " + $Status) -PercentComplete $PercentComplete
    }
}

# Just pulled this in, needs work
function Audit-deployments () {
    param(
        [Parameter(Mandatory = $False)][String]$outputpath
    )
	
    $prep = "\\generalvs\general\ITServices\Department\Records\Accreditation\current-semester-prep"
    $allresults = @()
    write-verbose "Getting all Managed Lab device collections from SCCM"
    $labs = (Get-CfgCollectionsByFolder "Managed Labs")

    write-verbose "Gathering info for each lab"
    $Progress = [ProgressBar]::New("Updating lab info", $labs.count)
    foreach ($lab in $labs) {
        $Progress.Advance($Lab.CollectionName)
        $LabMembers = Get-CfgCollectionMembers $Lab.CollectionName
        $labresult = $LabMembers |
            get-pcinfo -properties OSBuild, InstallDate, LastActiveDate, LastBoot, WarrantyEndDate, Model |
            add-member Collection $Lab.CollectionName -passthru
        $labresult | export-csv -path "$($lab.CollectionName).csv"
        $allresults = $allresults + $labresult
        # end result
    }
}
Function Convert-ToDate($MgmtDateTime) {
    [management.managementDateTimeConverter]::ToDateTime($MgmtDateTime)
}
Function Extract-OU($DistinguishedName, [Switch]$Short) {
    if ($Short) {
        $OU = ""
        $OUParts = ($DistinguishedName | select-string -allmatches "(?<OU>OU=[^,]*)").matches.value
        foreach ($OUPart in $OUParts) {
            if ($OUPart -eq "OU=Workstations") {
                continue
            }
            if ($OU) {
                $OU = ($OUPart -replace "OU=") + "/" + $OU
            } else {
                $OU = ($OUPart -replace "OU=")
            }
        }
    } else {
        $OU = $DistinguishedName -replace "(?<CN>CN=[^,]*,)"
    }
    return $OU
}

Function Get-LabInfo ($Collection) {

    $Query=@"
        SELECT Sys.*,Os.*,Comp.*,Client.*
        FROM SMS_R_System Sys
        JOIN SMS_G_System_OPERATING_SYSTEM Os ON Sys.ResourceId = Os.ResourceID
        JOIN SMS_G_System_COMPUTER_SYSTEM Comp ON Sys.ResourceID = Comp.ResourceID
        JOIN SMS_G_System_CH_ClientSummary Client ON Sys.ResourceID = Client.ResourceID
        WHERE Sys.ResourceId in
        (
            SELECT ResourceID
            FROM SMS_FullCollectionMembership
            JOIN SMS_Collection
            ON SMS_FullCollectionMembership.CollectionID = SMS_Collection.CollectionID
            WHERE SMS_Collection.name LIKE '$Collection'
        )
"@

    $QueryResult = Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_SC1" -Query $Query

    $NiceResult = $QueryResult | Select @{Name="Collection";Expression={($Collection)}},
        @{Name="ComputerName";Expression={$_.Sys."Name"}},
        @{Name="OU";Expression={ Extract-OU -short $_.Sys."DistinguishedName" }},
        @{Name="InstallDate";Expression={Convert-ToDate $_.Os."InstallDate"}},
        @{Name="OSName";Expression={$_.Os."Caption"}},
        @{Name="OSBuild";Expression={$_.Os."BuildNumber"}},
        @{Name="OSVersion";Expression={$_.Os."Version"}},
        @{Name="LastBoot";Expression={Convert-ToDate $_.Os."LastBootUpTime"}},
        @{Name="LastLogon";Expression={Convert-ToDate $_.Sys."LastLogonTimestamp"}},
        @{Name="LastPolicyRequest";Expression={Convert-ToDate $_.Client."LastPolicyRequest"}},
        @{Name="LastActiveDate";Expression={Convert-ToDate $_.Client."LastActiveTime"}},
        @{Name="MacAddresses";Expression={$_.Sys."MacAddresses"}},
        @{Name="Manufacturer";Expression={$_.Comp."Manufacturer"}},
        @{Name="Model";Expression={$_.Comp."Model"}},
        @{Name="User";Expression={$_.Comp."UserName"}},
        @{Name="Memory";Expression={$_.Os."TotalVisibleMemorySize"}},
        @{Name="VirtualMemory";Expression={$_.Os."TotalVirtualMemorySize"}},
        @{Name="IPSubnets";Expression={$_.Sys."IPSubnets"}},
        @{Name="IPAddresses";Expression={$_.Sys."IPAddresses"}},
        @{Name="ADGroups";Expression={$_.Sys."SecurityGroupName"}}

    $defaultProperties = @("Collection","ComputerName","OU","InstallDate","OSBuild","LastBoot","LastActiveDate","Model")
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $NiceResult | Add-Member MemberSet PSStandardMembers $PSStandardMembers

    Return [PCInfo[]]$NiceResult
}
Export-ModuleMember -function Get-LabInfo

Function Get-LabSummary {
    $Labs = Import-CliXML $LabsCache
    $GroupedReport = Import-CliXML $LabReportCache | group collection -ashashtable -asstring

    $Summary = foreach ($LabName in $Labs.CollectionName) {
        $LabReport = $GroupedReport.$Labname
        $TotalPCs = $LabReport.Count
        $Win7Count = ($LabReport | ? { $_.OSBuild -eq "7601"}).OSBuild.Count
        $Down7Days = ($LabReport | ? {
                (-not $_.LastActiveDate) -or
                (new-timespan -start ($_.LastActiveDate) -end (get-date)) -gt (new-timespan -days 7) 
            }
        ).ComputerName.Count
        [LabSummary]@{
            Collection = $LabName
            Down7Days = $Down7Days
            Win7 = $Win7Count
            TotalPCs = $TotalPCs
        }
    }
    $Summary  | sort -property Win7
}
Export-ModuleMember -function Get-LabSummary
Set-Alias Summary Get-LabSummary
Export-ModuleMember -alias Summary

Function Get-LabReport ($substr) {
    $Labs = Import-CliXML $LabsCache

    if ($substr) {
        $GroupedReport = Import-CliXML $LabReportCache | group collection -ashashtable -asstring

        $CurrentLabs = $Labs.CollectionName | ? { $_ -like "*$($substr)*"}
        if ($CurrentLabs.Count -eq 0) {
            write-warning "No collections matched $substr"
            return
        }
        $Out = @()
        foreach ($Lab in $CurrentLabs) {
            $Out = $Out + $GroupedReport.$Lab
        }
    } else {
        #$Out = Import-XML $LabReportCache
        $Out = Import-CliXML $LabReportCache
    }

    $DefaultProperties = @("ComputerName","OU","OSBuild","InstallDate","LastBoot","LastActiveDate","Model","Collection")
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $Out | Add-Member MemberSet PSStandardMembers $PSStandardMembers -passthru
}
Export-ModuleMember -function Get-LabReport
Set-Alias Report Get-LabReport
Export-ModuleMember -alias Report

Function Publish-LabReport {
    $Labs = Import-CliXML $LabsCache
    $Report = Import-CliXML $LabReportCache |
        select @{Name="PC";Expression={$_.ComputerName}},
            OU,
            @{Name="Install Date";Expression={ "{0:dd/MM/yy}" -f $_.InstallDate }},
            OSBuild,
            @{Name="Last Boot";Expression={ "{0:dd/MM/yy}" -f $_.LastBoot }},
            @{Name="Last Logon";Expression={ "{0:dd/MM/yy}" -f $_.LastLogon }},
            @{Name="Last Policy";Expression={ "{0:dd/MM/yy}" -f $_.LastPolicyRequest }},
            @{Name="Last Active";Expression={ "{0:dd/MM/yy}" -f $_.LastActiveDate }},
            MacAddresses,
            Model,
            Collection
    $GroupedReport = $Report | group collection -ashashtable -asstring

    remove-item $reportdir\*.*

    $OutputPath = "$($reportdir)\\summary.html"
    summary | convertto-html > $OutputPath
    Add-HTMLHeaderExtras $OutputPath "USC Labs: Summary View"

    foreach ($Lab in $Labs.CollectionName) {
        $OutputPath = "$($reportdir)\\$($Lab).html"
        $GroupedReport.$Lab | select -exclude Collection | convertto-html > $OutputPath
        Add-HTMLHeaderExtras $OutputPath $Lab
    }
}
Export-ModuleMember -function Publish-LabReport

Function Add-HTMLHeaderExtras2 ($filename) {
    $ID = ...

    [xml]$plist = get-content $filename
    $node = $plist.SelectSingleNode("//array/string[text()='ABCD']")

    $newNode = $plist.CreateElement('string')
    $newNode.AppendChild($plist.CreateTextNode($ID)) >$null

$node.ParentNode.InsertAfter($newNode, $node)
}

Function Add-HTMLHeaderExtras ($filename, $Heading) {
    $headbit = @"

    <link rel="stylesheet" type="text/css" href="../css/labreport.css">
    <link rel="stylesheet" type="text/css" href="../css/tablesort.css">
    <script src='../js/tablesort/tablesort.js' type="text/javascript"></script>
    <script src='../js/tablesort/sorts/tablesort.number.js' type="text/javascript"></script>
    <script src='../js/tablesort/sorts/tablesort.date.js' type="text/javascript"></script>
    <script src='../js/labreport.js' type="text/javascript"></script>
"@
    $Heading = if ($Heading) {
        "<h1>$Heading</h1>`n"
    } else {
        ""
    }
    $aftertablebit = @"
    <script>
        init_rowevents();
        new Tablesort(document.getElementById('labreport'));
    </script>
"@
    (Get-Content $filename).
        replace('<head>','<head>'+$headbit).
        replace('<body>','<body>'+$Heading).
        replace('</table>','</table>'+$aftertablebit).
        replace('<table>','<table id="labreport">').
        replace('<tr><th>',"<thead>`n<tr><th>").
        replace('</th></tr>',"</th></tr>`n</thead>") |
        Set-Content $filename
    .\html\summary.html
}

# Main

Update-CachedDataIfNecessary

#-------------------------------------------------------------------------------------------

#Not working yet


#Not working yet
# Format dates as yyyy-MM-dd
$currentThread = [System.Threading.Thread]::CurrentThread
$culture = [CultureInfo]::InvariantCulture.Clone()
$culture.DateTimeFormat.ShortDatePattern = 'yyyy-MM-dd'
$currentThread.CurrentCulture = $culture
$currentThread.CurrentUICulture = $culture