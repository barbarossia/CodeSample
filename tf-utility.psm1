$vsPath = “C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer”
Add-Type -Path “$vsPath\Microsoft.TeamFoundation.Client.dll”
#Add-Type -Path "$vs2015Path\Microsoft.TeamFoundation.WorkItemTracking.Client"
#Add-Type -Path “$vs2015Path\Microsoft.TeamFoundation.Common.Client.dll”
Add-Type -Path “$vsPath\Microsoft.TeamFoundation.VersionControl.Client.dll”
Add-Type -Path “$vsPath\Microsoft.TeamFoundation.VersionControl.Common.dll”

[string]$rootPath = "D:\Work\BackOffice"
[string]$tfsUrl = “http://vs.mycorp.com/tfs/DefaultCollection” 
[string]$devPath = "$rootPath\DEV\Iteration"
[string]$mainPath = "C:\MAIN"
[string]$releasePath = "C:\Release"
[string]$rootTFS = "$/BackOffice"
[string]$devTFS = "$rootTFS/DEV/Iteration"
[string]$mainTFS = "$rootTFS/MAIN"
[string]$releaseTFS = "$rootTFS/RELEASE"

$tf = $vsPath + "\tf.exe"


function add-file([string] $file){
&$tf add $file | Out-Null
}
function checkout-file([string] $file)
{
    "Checking out $file"
    <#
    if ((Get-PSSnapin -Name  Microsoft.TeamFoundation.PowerShell -ErrorAction SilentlyContinue) -eq $null )
    {
        Add-PsSnapin  Microsoft.TeamFoundation.PowerShell -ErrorAction SilentlyContinue

        if ((Get-PSSnapin -Name  Microsoft.TeamFoundation.PowerShell -ErrorAction SilentlyContinue) -eq $null )
        {
            #try to check out the code using command line tf.exe
            &"C:\Program Files (x86)\Microsoft Visual Studio 11.0\Common7\IDE\TF.exe" checkout $file | Out-Null
        }
        else{
            #checkout the file using snapin
            Add-TfsPendingChange -Edit $file | Out-Null
        }
    }else{
        #checkout the file using snapin
        Add-TfsPendingChange -Edit $file | Out-Null
    }
    #>
    &$tf checkout $file | Out-Null
}
<#
.Synopsis
   Gets a list of all changesets from $Source that have not been merged into $Destination.
.DESCRIPTION
   Gets a list of all changesets from $Source that have not been merged into $Destination.
.EXAMPLE
   $tfs = Get-TfsServer http://<hostname>:8080/tfs
   Get-MergeCandidates -Server $tfs -Source "$/Some/Source/Branch" -Destination "$/Some/Destinsation/Branch"
   The Get-TfsServer cmdlet is included as part of the PowerShell snapin provided by the TFS Power Tools 
   extension for Visual Studio. If everything is installed correctly, you should be able to load the snapin
   with the following command:
   Add-PSSnapin Microsoft.TeamFoundation.PowerShell
#>
function Get-TfsMergeCandidates {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [Microsoft.TeamFoundation.Client.TfsTeamProjectCollection]
        $Server,

        [Parameter(Mandatory = $true)]
        [string]
        $Source,
        
        [Parameter(Mandatory = $true)]
        [string]
        $Destination
        
    )

    $service = $Server.GetService([Microsoft.TeamFoundation.VersionControl.Client.VersionControlServer])

    $candidates = $service.GetMergeCandidates(
        $Source, 
        $Destination, 
        [Microsoft.TeamFoundation.VersionControl.Client.RecursionType]::Full
    )

    $candidates | select -ExpandProperty changeset
}



function Get-CheckIn {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]
        $tfsURL,

        [Parameter(Mandatory = $true)]
        [string]
        $LocalPath,
        
        [Parameter(Mandatory = $false)]
        [string]
        $workitemId,
                
        [Parameter(Mandatory = $false)]
        [string]$checkInComments = "Checked in from PowerShell"
        
    )
    #[psobject] $tfs=[Microsoft.TeamFoundation.Client.TeamFoundationServerFactory]::GetServer($tfsURL)
    $tfs = [Microsoft.TeamFoundation.Client.TfsTeamProjectCollectionFactory]::GetTeamProjectCollection($tfsURL)
$versionControlType = [Microsoft.TeamFoundation.VersionControl.Client.VersionControlServer] 
$versionControlServer = $tfs.GetService($versionControlType) 
$WorkstationType = [Microsoft.TeamFoundation.VersionControl.Client.Workstation]
$WorkspaceInfo = $WorkstationType::Current.GetLocalWorkspaceInfo($LocalPath)
$Collection = [Microsoft.TeamFoundation.Client.TfsTeamProjectCollectionFactory]::GetTeamProjectCollection($WorkspaceInfo.ServerUri)
$Collection.EnsureAuthenticated()

$workspace =  $WorkspaceInfo.GetWorkspace($Collection)
#$pendingChanges = $workspace.GetPendingChanges();

#$wici = New-Object Microsoft.TeamFoundation.VersionControl.Client.WorkItemCheckinInfo($workItem, [Microsoft.TeamFoundation.VersionControl.Client.WorkItemCheckinAction]::Associate);
#$workItems = @($wici)
$changesetId = $workspace.CheckIn($workspace.GetPendingChanges(), $checkInComments)


if ($workitemId) {
    $wit=$tfs.Getservice([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])
    $item=$wit.GetWorkItem($workitemId)
    $ch = $versionControlServer.GetChangeset($changesetId);
   
    
    $linkType = $wit.RegisteredLinkTypes["Fixed in Changeset"];
    $witLink = New-Object Microsoft.TeamFoundation.WorkItemTracking.Client.ExternalLink($linkType, $ch.ArtifactUri.AbsoluteUri);
    $witLink.Comment = $checkInComments
    $item.Links.Add($witLink);
    $item.History = "Automatically associated with changeset " + $changesetId;
    
    $item.Save() | Out-Null; 
    
}

return $($ch.ChangesetId)

}

function Associate-Task{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]
        $tfsURL,

        
        [Parameter(Mandatory = $false)]
        [string]
        $sourceid,
        [Parameter(Mandatory = $false)]
        [string]
        $linkType='Related',
        [Parameter(Mandatory = $false)]
        [string]
        $targetid
                
        
    )
    #[psobject] $tfs=[Microsoft.TeamFoundation.Client.TeamFoundationServerFactory]::GetServer($tfsURL)
    $tfs = [Microsoft.TeamFoundation.Client.TfsTeamProjectCollectionFactory]::GetTeamProjectCollection($tfsURL)
    $wit=$tfs.Getservice([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])
    $source=$wit.GetWorkItem($sourceid)    
    $EndLinkType = $wit.WorkItemLinkTypes.LinkTypeEnds[$linkType];
    $link = New-Object Microsoft.TeamFoundation.WorkItemTracking.Client.RelatedLink($EndLinkType, $targetid)

    $source.Links.Add($link);
    $result = $source.Validate(); 
    $source.Save(); 


}

function Get-CheckInMain {
param(  

  [String] $workitemId = 306355,
  [String] $location = 'SAM-CAP'
) 

return Get-CheckIn -tfsURL $tfsUrl -LocalPath $mainPath -workitemId $workitemId -checkInComments "merge to main"

}


function Get-checkinRelease {
param(  

  [String] $workitemId = 306655,
  [string] $release = 2.51,
  [string] $location
) 

$current_releasePath = "$releasePath/$release"
return Get-Checkin -tfsURL $tfsUrl -LocalPath $current_releasePath -workitemId $workitemId -checkInComments "merge to release $release"

}

function Get-Merge {
param(  

  [String] $changeId,
  [string] $source,
  [string] $target
) 

#Set-Location $rootPath; 

& $tf merge $source $target /recursive /version:C$changeId~C$changeId
}


function Get-MergeMain {
param(  

  [String] $changeId
) 

Get-Merge -changeId $changeId -source $devTFS -target $mainTFS

}

function Get-MergeRelease {
param(  

  [String] $changeId,
  [string] $release,
  [string] $location
) 

$source = "$mainTFS/$location"
$target = "$releaseTFS/$release/MAIN/$location"

Get-Merge -changeId $changeId -source $source -target $target

}

function Get-MergeReleaseAPR {
param(  

  [String] $changeId,
  [string] $release,
  [string] $location
) 

$releaseAPR = $release + "APR"

$source = "$releaseTFS/$release/MAIN/$location"
$target = "$releaseTFS/$releaseAPR/MAIN/$location"

Get-Merge -changeId $changeId -source $source -target $target

}

function Pause($message = "Do you want to continue... (default is yes)") {
  #Write-Host -NoNewLine $title
  $message = $message + " Do you want to continue... (default is yes)"
  write-host $message -ForegroundColor Yellow
  
  $readHost = Read-Host "(y/n)"
  $result = 0
  switch($readHost){
    y { $result = 0 }
    n { $result = 1 }
    default {$result = 0 }
  }
  return $result
}


function Get-DropMerge {
param(  

  [String] $changeId
) 

$source = $devTFS
$target = $mainTFS


Set-Location $rootPath; 

& $tf merge $source $target /discard /recursive /version:C$changeId~C$changeId
}

function Get-MergeMainAll {
param(  
  [String] $changeId = 271922,
  [string] $workitemid = 307695,
  [string] $location = 'COnfig'
) 
    Get-MergeMain $changeId 
    $result = Pause
    if ($result -eq 0) {
        return Get-CheckInMain -workitemId $workitemid -location $location
    }
}

function Get-MergeReleaseAll {
param(  
  [String] $changeId = 271922,
  [string] $workitemid = 307695,
  [string] $release = 2.52,
  [string] $location = 'COnfig'
) 
    Get-MergeRelease $changeId -release $release -location $location
    $result = Pause
    if ($result -eq 0) {
        return Get-checkinRelease -workitemId $workitemid -release $release -location $location
    }
}

function Get-MergeReleaseAPRAll {
param(  
  [String] $changeId = 271922,
  [string] $workitemid = 307695,
  [string] $release = '2.52',
  [string] $location = 'COnfig'
) 
    Get-MergeReleaseAPR -changeId $changeId -release $release -location $location
    $result = Pause
    if ($result -eq 0) {
        $releaseAPR = $release + "APR"
        $result = Get-checkinRelease -workitemId $workitemid -release $releaseAPR -location $location
        return $result
    }
}


function Get-WorkItem {
 [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]
        $tfsURL,
        
        [Parameter(Mandatory = $false)]
        [string]
        $workitemId
                        
    )

    $tfs = [Microsoft.TeamFoundation.Client.TfsTeamProjectCollectionFactory]::GetTeamProjectCollection($tfsURL)
    $wit=$tfs.Getservice([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])
    $item = $wit.GetWorkItem($workitemId)
    return $item
}


function New-WorkItem()
{
 [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]
        $tfsURL,
        
        [Parameter(Mandatory = $false)]
        [string]
        $title,
        [Parameter(Mandatory = $false)]
        [string]
        $State = 'Proposed',
        [Parameter(Mandatory = $false)]
        [int]
        $Estimate,
        [Parameter(Mandatory = $false)]
        [string]
        $Assigned,     
        [Parameter(Mandatory = $false)]
        [int]
        $sourceid                    
    )
    # These *should* be registered in the GAC.
    # The version numbers will likely have to change as new versions of Visual Studio are installed on the server.



    $tfs = [Microsoft.TeamFoundation.Client.TfsTeamProjectCollectionFactory]::GetTeamProjectCollection($tfsURL)
    $type = [Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore]
    $store = [Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore] $tfs.GetService($type)
    $project = $store.Projects["BackOffice"]
    $workItem = New-Object Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem($project.WorkItemTypes['Task'])
    $workItem.Title = $title
    $workitem.Fields['State'].Value = $State
    $workitem.Fields['Original Estimate'].Value = $Estimate
    $workitem.Fields['Remaining Work'].Value = 0
    $workitem.Fields['Remaining Work'].Value = 0
    $workitem.Fields['Assigned To'].Value = $Assigned
    #$workitem.Fields["System.State"].Value = "Proposed"
    #$workitem.IterationPath = 'BackOffice\2019\01 - January\AP Support'
    #$workitem.IterationId = 7213
    #$workitem.NodeName = 'BackOffice'
    #$workitem.AreaPath = 'BackOffice'
    #$workitem.IterationPath = 'BackOffice\2019\01 - January\AP Support'

    $source =  Get-WorkItem -tfsurl $tfsUrl -workitemid $sourceid

    $workitem.IterationPath = $source.IterationPath
    $workitem.AreaPath = $source.AreaPath
    
    $result = $workitem.Validate();
    $workitem.Save()
    
    Associate-Task -tfsURL $tfsUrl -sourceid $sourceid -targetid $workitem.Id -linkType 'Child'
    return $workitem
}

function Format-Xml {
<#
.SYNOPSIS
Format the incoming object as the text of an XML document.
#>
    param(
        ## Text of an XML document.
        [Parameter(ValueFromPipeline = $true)]
        [string[]]$Text
    )

    begin {
        $data = New-Object System.Collections.ArrayList
    }
    process {
        [void] $data.Add($Text -join "`n")
    }
    end {
        $doc=New-Object System.Xml.XmlDataDocument
        $doc.LoadXml($data -join "`n")
        $sw=New-Object System.Io.Stringwriter
        <#
        #Settings object will instruct how the xml elements are written to the file
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Indent = $true
        #NewLineChars will affect all newlines
        $settings.NewLineChars ="`r`n"
        #Set an optional encoding, UTF-8 is the most used (without BOM)
        $settings.Encoding = New-Object System.Text.UTF8Encoding( $false )
        #>
        #$encoding = [System.Text.Encoding]::UTF8
        $writer=New-Object System.Xml.XmlTextWriter($sw)
        $writer.Formatting = [System.Xml.Formatting]::Indented

        $doc.WriteContentTo($writer)
        $sw.ToString()
        $writer.Flush()
        $writer.Close()
    }
}

function add-script{
param( [string]$script)
[System.XML.XMLElement]$build = $xdoc.CreateElement("Build", $ns)

$build.SetAttribute("Include", $script)
[System.XML.XMLElement]$SubType = $xdoc.CreateElement("SubType", $ns)
$SubType.InnerText = 'Code'
$build.AppendChild($SubType)

$sqlitems.AppendChild($build)
}

function get-workitems {

$tfs = [Microsoft.TeamFoundation.Client.TfsTeamProjectCollectionFactory]::GetTeamProjectCollection($tfsURL)
$wit=$tfs.Getservice([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])

$StoredQuery = @"
SELECT [System.ID], 
[System.Title], 
[System.Description], 
[System.State] 
FROM WorkItems 
WHERE [System.TeamProject] = 'BackOffice' 
AND [System.AssignedTo] = @Me

AND System.ChangedDate >= '8/1/2019'
ORDER BY System.ChangedDate desc
"@
 #and [System.AssignedTo] = 'CHN\ZhangBobo1'
 #-- [System.TeamProject] = 'BackOffice' 
 #AND [System.State] = 'Proposed'
$workItems = $wit.Query($StoredQuery)
 
foreach($item in $workItems)
{
write-output $item['System.Title']
}
}


function tfs-updateSource(){

 
 
cls;

#Add-PSSnapin Microsoft.TeamFoundation.PowerShell 

Get-TfsWorkspace -Path $devPath

Update-TfsWorkspace -Recurse $devPath | Out-GridView

 

Write-Host ""

$now = Get-Date

Write-Host "As of: " $now
}

function tfs-get {
& $tf get /recursive $devTFS $devPath
}

Export-ModuleMember -Variable tfsUrl, devPath, mainPath, devTFS, mainTFS, releaseTFS
Export-ModuleMember -Function Get-TfsMergeCandidates `
                            , Get-CheckIn, Get-CheckInMain, Get-MergeMain, Get-CheckinRelease `
                            , Get-MergeRelease, Get-MergeReleaseAPR, Get-DropMerge, Get-MergeMainAll `
                            , Get-MergeReleaseAll, Get-MergeReleaseAPRAll, Get-WorkItem `
                            , Associate-Task `
                            , New-WorkItem  `
                            , checkout-file `
                            , add-file `
                            , Format-Xml `
                            , add-script `
                            , get-workitems `
                            , tfs-updateSource `
                            , tfs-get