function Install-Az {
	# If you don't have "az" utility, install latest one
	Invoke-WebRequest -Uri https://aka.ms/installazurecliwindows -OutFile .\AzureCLI.msi; Start-Process msiexec.exe -Wait -ArgumentList '/I AzureCLI.msi /quiet'; rm .\AzureCLI.msi
}


function Reverse
{ 
 $arr = @($input)
 [array]::reverse($arr)
 $arr
}

function NormalizeDateTime([string]$dateTime){
	if ([string]::IsNullOrEmpty($dateTime)){
		return $null;
	}
	
	return [System.DateTime]::Parse($dateTime).ToUniversalTime().ToString("s")
}

Class DevOpsConfig {
	# Project config
	[string]$org
	[string]$project
	[string]$team
	[string]$areaPath

	# Time window before and after sprint start which is OK for sprint modification.
	[int]$planningDaysBefore=4
	[int]$planningDaysAfter=2
	
	DevOpsConfig([string]$org,[string]$project,[string]$team,[string]$areaPath){
		$this.org=$org
		$this.project=$project
		$this.team=$team
		$this.areaPath=$areaPath
	}   
}      

function New-DevOpsConfig(
	[Parameter(Mandatory=$True)][string]$org,
	[Parameter(Mandatory=$True)][string]$project,
	[Parameter(Mandatory=$True)][string]$team,
	[Parameter(Mandatory=$True)][string]$areaPath,
	[int]$planningDaysBefore=4,
	[int]$planningDaysAfter=2){
	$config = [DevOpsConfig]::New($org,$project,$team,$areaPath)
	$config.planningDaysBefore=$planningDaysBefore
	$config.planningDaysAfter=$planningDaysAfter
	
	return $config
} 

function Get-WorkItemHistory([Parameter(Mandatory=$True)][DevOpsConfig]$config, [Parameter(Mandatory=$True)][string]$id)
{
	$revisions = (az devops invoke --http-method GET --resource revisions --area wit --organization $config.org --api-version 5.1-preview --route-parameters id="$id" -o json | ConvertFrom-Json).value
	
	$current = ($revisions | measure -Maximum -Property rev).Maximum
	$result = $revisions | where {$_.rev -eq $current}
	
	Add-Member -InputObject $result -MemberType NoteProperty -Name "revisions" -Value $revisions
	Add-Member -InputObject $result -MemberType NoteProperty -Name "current" -Value $result
	
	return $result
}

Class DevOpsSprintSet {
	[Object[]]$sprints
	[Object]$current
	[Object]$last
	[Object]$next
	
	DevOpsSprintSet([Object[]]$sprints){
		$this.sprints=$sprints
		$this.current = $sprints | where {$_.timeframe -eq "current"} | select -first 1
		$this.last=$this.GetPrevious($this.current)
		$this.next=$this.GetNext($this.current)
	}
	
	[Object]GetSprintByName([string]$name){
		return $this.sprints | where {$_.name -ieq $name} | select -First 1
	}
	
	[Object]GetSprintByNumber([int]$number){
		return $this.GetSprintByName("Sprint $number")
	}
	
	[Object]GetSprintByPath([string]$path){
		return $this.sprints | where {$_.path -ieq $path} | select -First 1
	}	
	
	[Object]GetNext([object]$sprint){
		$sprint=$this.GetSprintByPath($sprint.path)
		$index=$this.sprints.IndexOf($sprint)
		
		if ($index -lt 0 -or $index -ge ($this.sprints.count-1)) {
			return $null
		}
		
		return $this.sprints[$index+1]
	}	
	
	[Object]GetPrevious([object]$sprint){
		$sprint=$this.GetSprintByPath($sprint.path)
		$index=$this.sprints.IndexOf($sprint)
		
		if ($index -le 0 -or $index -ge $this.sprints.count) {
			return $null
		}
		
		return $this.sprints[$index-1]
	}
}

function Get-SprintSet([Parameter(Mandatory=$True)][DevOpsConfig]$config) {
	# This is only required because of a bug in current version (color of console permanently changed to invisible...logging fixes it, but you ahve to throw out logs...)
	$json =  ((az boards iteration  team list --team $config.team -p $config.project  --org $config.org -o json --debug) 2>$null)

	# Select last completed sprint, or you can select any previous sprint by id
	$sprints = ($json | ConvertFrom-Json) | select name, path -expand attributes | foreach {
		$_planStart=[System.DateTime]::Parse($_.startDate).ToUniversalTime().Subtract([System.TimeSpan]::FromDays($config.planningDaysBefore)).ToString("s")
		$_planEnd=[System.DateTime]::Parse($_.startDate).ToUniversalTime().Add([System.TimeSpan]::FromDays($config.planningDaysAfter)).ToString("s")
		Add-Member -InputObject $_ -MemberType NoteProperty -TypeName "String" -Name "planStartDate" -Value $_planStart
		Add-Member -InputObject $_ -MemberType NoteProperty -TypeName "String" -Name "planEndDate" -Value $_planEnd
		Add-Member -InputObject $_ -MemberType AliasProperty -TypeName "String" -Name "endDate" -Value finishDate
		$_
	}
	
	return [DevOpsSprintSet]::New($sprints)
}

function Get-SprintWorkItems(
	[Parameter(Mandatory=$true)][DevOpsConfig]$config,
	[Parameter(Mandatory=$true)]$sprint,
	[Parameter(Mandatory=$false)][switch]$idOnly,
	[Parameter(Mandatory=$false)][switch]$includeHistory,
	[Parameter(Mandatory=$false)][switch]$includeRelated) {
	# Items ever marked as sprint members
	$sprintItems =  az boards query --org $config.org -o json --wiql "SELECT [Id],[Title],[Work Item Type],[State],[Changed Date],[Created Date],[Activated Date],[Closed Date],[Resolved Date] FROM workitems WHERE (EVER [System.IterationPath] = '$($sprint.path)') AND [System.AreaPath] UNDER '$($config.areaPath)'" | ConvertFrom-Json

	if ($includeRelated){
		# Items started before the sprint and finished after sprint start
		$sprintItems +=  az boards query --org $org -o json --wiql "SELECT [Id],[Title],[Work Item Type],[State],[Changed Date],[Created Date],[Activated Date],[Closed Date],[Resolved Date] FROM workitems WHERE [Created Date] < '$($sprint.startDate)' AND [Activated Date] <> ''  AND [Activated Date] < '$($sprint.startDate)' AND [Changed Date] > '$($sprint.startDate)' AND [State] in ('Resolved','Closed','Removed') AND [System.AreaPath] UNDER '$($config.areaPath)'" | ConvertFrom-Json

		# Items started before the sprint but still not finished
		$sprintItems +=  az boards query --org $org -o json --wiql "SELECT [Id],[Title],[Work Item Type],[State],[Changed Date],[Created Date],[Activated Date],[Closed Date],[Resolved Date] FROM workitems WHERE [Created Date] < '$($sprint.startDate)' AND [Activated Date] <> '' AND [Activated Date] < '$($sprint.startDate)' AND [State] = 'Active' AND [System.AreaPath] UNDER '$($config.areaPath)'" | ConvertFrom-Json

		# Items started in the sprint
		$sprintItems +=  az boards query --org $org -o json --wiql "SELECT [Id],[Title],[Work Item Type],[State],[Changed Date],[Created Date],[Activated Date],[Closed Date],[Resolved Date] FROM workitems WHERE [Activated Date] <> '' AND [Activated Date] >= '$($sprint.startDate)' AND [Activated Date] <= '$($sprint.endDate)' AND [System.AreaPath] UNDER '$($config.areaPath)'" | ConvertFrom-Json
	}
	
	$sprintItems=$sprintItems | Sort-Object -Property Id -Unique
	$sprintItemIds=$sprintItems | select -ExpandProperty Id
	
	if ($idOnly) {
		return $sprintItemIds
	}
	
	if ($includeHistory -and $sprintItems.count -gt 0) {
		#init
		$i=0; $sprintItemsWithHistory = @{};
		Write-Progress -Activity "Querying work item history" -status "Querying workitem 1/$($sprintItems.count)" -percentComplete 0
		
		# Now you can get a coffee, as it is 2-3 sec per item
		$i=1; $sprintItemsWithHistory = @{};
		$sprintItemIds | foreach { 
				$sprintItemsWithHistory.Add($_,(Get-WorkItemHistory -config $config -id $_ ));
				$i++;
				Write-Progress -Activity "Querying work item history" -status "Querying workitem $i/$($sprintItems.count)" -percentComplete (($i-1)/$sprintItems.count*100)
			}
			
		Write-Progress -Activity "Querying work item history" -status "Done" -percentComplete 100
		
		$sprintItems=@($sprintItemsWithHistory.Values)
	}
	 
	return $sprintItems
}

#
# Helpers to investigate items
#
function Get-RevisionAt([object]$workItem, [string]$dateTime)
{
	$dateTime = NormalizeDateTime $dateTime
	
	$sprintRev=$null
	$workItem.revisions | reverse | foreach {
		if (($sprintRev -eq $null) -and ($_.fields."System.ChangedDate" -lt $dateTime)) {
			$sprintRev=$_
		}
	}
	
	return $sprintRev
}

# Helper to check whether an item was int he sprint 
function Check-SprintAt([object]$workItem, [object]$sprint, [string]$dateTime)
{
	$dateTime = NormalizeDateTime $dateTime
	$sprintRev = Get-RevisionAt $workItem $dateTime
	
	return ($sprintRev -ne $null) -and ($sprintRev.fields."System.IterationPath" -eq $sprint)
}


function Get-NewRevisionsBetween([object]$workItem, [string]$startDateTime, [string]$endDateTime)
{
	if ($startDateTime -eq $null -or $endDateTime -eq $null){
		return $null
	}
	
	$startDateTime = NormalizeDateTime $startDateTime
	$endDateTime = NormalizeDateTime $endDateTime
	echo $startDateTime
	echo $endDateTime	
	$revisions=$workItem.revisions | where  {
	($_.fields."System.ChangedDate" -gt $startDateTime) -and  ($_.fields."System.ChangedDate" -lt $endDateTime)}
	echo $revisions
	return $revisions
}

function Check-ActivationBefore([object]$workItem, [string]$dateTime)
{
	$dateTime = NormalizeDateTime $dateTime
	
	if ($dateTime -eq $null){
		return $false;
	}
	
	$activationDate=$workItem.fields."Microsoft.VSTS.Common.ActivatedDate"
	
	return  ($activationDate -ne $null) -and ($activationDate -lt $dateTime)
}


function Check-ActivationAfter([object]$workItem, [string]$dateTime) {
	$dateTime = NormalizeDateTime $dateTime
	
	if ($dateTime -eq $null){
		return $false;
	}
	
	$activationDate=$workItem.fields."Microsoft.VSTS.Common.ActivatedDate"
	
	return  ($activationDate -ne $null) -and ($activationDate -gt $dateTime)
}

function Check-Done([object]$workItem){
	return ($workItem.fields."System.State" -eq "Resolved") -or ($workItem.fields."System.State" -eq "Closed") -or ($workItem.fields."Microsoft.VSTS.Common.ResolvedDate" -ne $null) -or ($workItem.fields."Microsoft.VSTS.Common.ClosedDate" -ne $null)
}

function Get-DoneDate([object]$workItem){
	if (-not (Check-Done $workItem)){
		return $null
	}
	
	$resolvedDate = $workItem.fields."Microsoft.VSTS.Common.ResolvedDate"
	$closedDate = $workItem.fields."Microsoft.VSTS.Common.ClosedDate"
	$date = if ($resolvedDate -ne $null) {$resolvedDate;} elseif ($closedDate -ne $null){$closedDate;} else {null;}

	return $date;
}

function Get-DoneBy([object]$workItem){
	if (-not (Check-Done $workItem)){
		return $null
	}
	
	$resolvedBy = $workitem.fields."Microsoft.VSTS.Common.ResolvedBy".uniqueName
	$closedBy = $workItem.fields."Microsoft.VSTS.Common.ClosedBy".uniqueName
	$ret = if ($resolvedBy -ne $null) {$resolvedBy;} elseif ($closedBy -ne $null){$closedBy;} else {$null;}
	
	return $ret;
}

function Check-DoneBefore([object]$workItem, [string]$dateTime) {
	$dateTime = NormalizeDateTime $dateTime
		
	if ($dateTime -eq $null){
		return $false;
	}
		
	$date = Get-DoneDate $workItem
	
	return (($date -ne $null) -and ($date -lt $dateTime))
}

function Check-DoneAfter([object]$workItem, [string]$dateTime) {
	$dateTime = NormalizeDateTime $dateTime
		
	if ($dateTime -eq $null){
		return $false;
	}
		
	$date = Get-DoneDate $workItem
	
	return (($date -ne $null) -and ($date -gt $dateTime))
}

function Check-Planned([object]$workItem, [object]$sprint){
	return Check-SprintAt $workItem $sprint $sprint.planEndDate
}

function Check-Unplanned([object]$workItem, [object]$sprint) {
	 $planned=Check-Planned $workItem $sprint
	 $addedLater=@(Get-NewRevisionsBetween $workItem $sprint.planEndDate $sprint.endDate | where {$_.fields."System.IterationPath" -eq $sprint.path}).Count -gt 0
	 
	 return (-not $planned) -and $addedLater
}

function Check-Removed([object]$workItem, [object]$sprint) {
	$planned=Check-Planned $workItem $sprint
	$addedLater=@(Get-NewRevisionsBetween $workItem $sprint.planEndDate $sprint.endDate | where {$_.fields."System.IterationPath" -eq $sprint.path}).Count -gt 0
	$removed=(-not (Check-SprintAt $workItem $sprint $sprint.endDate))
	
	return ($planned -or $addedLater) -and $removed
}

function Check-EverAssigned([object]$workItem, [object]$sprint){
	$ret=$false
	$workItem.revisions | foreach {if ($_.Fields."System.IterationPath" -eq $sprint.path) {$ret=$true} }
	
	return $ret
}

function Get-FirstAssignedDate([object]$workItem, [object]$sprint){
	$ret=$null
	$workItem.revisions | foreach {if (($ret -eq $null) -and ($_.Fields."System.IterationPath" -eq $sprint.path)) {$ret=$_.fields."System.ChangedDate"} }
	
	return $ret
}

function Get-FirstAssignedBy([object]$workItem, [object]$sprint){
	$ret=$null
	$workItem.revisions | foreach {if (($ret -eq $null) -and ($_.Fields."System.IterationPath" -eq $sprint.path)) {$ret=$_.fields."System.ChangedBy"} }
	
	return $ret.uniqueName
}

function Get-SprintReport([object]$sprint, [object[]]$workItems, [string]$path){
	$items=$workItems | foreach {
		if ($_.current -eq $null){
			throw "Riport can only be made from sprint items with revision history"
		}
	
		$item=@{}
		$item.id=$_.id
		$item.title=$_.fields."System.Title"
		$item.type=$_.fields."System.WorkItemType"
		$item.state=$_.fields."System.State"
		$item.revisions=$_.revisions.count
		$item.sprint=$_.fields."System.IterationPath"
		$item.assigned_date=Get-FirstAssignedDate $_ $sprint
		$item.assigned_by=Get-FirstAssignedBy $_ $sprint
		$item.ever_assigned=Check-EverAssigned $_ $sprint
		$item.created_date=$_.fields."System.CreatedDate"
		$item.created_by=$_.fields."System.CreatedBy".uniqueName
		$item.activated_date=$_.fields."Microsoft.VSTS.Common.ActivatedDate"
		$item.activated_by=$_.fields."Microsoft.VSTS.Common.ActivatedBy".uniqueName
		$item.resolved_date=$_.fields."Microsoft.VSTS.Common.ResolvedDate"
		$item.resolved_by=$_.fields."Microsoft.VSTS.Common.ResolvedBy".uniqueName
		$item.closed_date=$_.fields."Microsoft.VSTS.Common.ClosedDate"
		$item.closed_by=$_.fields."Microsoft.VSTS.Common.ClosedBy".uniqueName
		$item.done_date=Get-DoneDate $_
		$item.done_by=Get-DoneBy $_
		$item.changed_date=$_.fields."System.ChangedDate"
		$item.change_by=$_.fields."System.ChangedBy".uniqueName
		$item.planned=Check-Planned $_ $sprint
		$item.removed=Check-Removed $_ $sprint
		
		if (Check-ActivationBefore $_ $sprint.startDate){
			$item.started="BeforeSprint"
		}
		elseif (Check-ActivationAfter $_ $sprint.endDate){
			$item.started="AfterSprint"
		}
		elseif ((Check-ActivationAfter $_ $sprint.startDate) -and (Check-ActivationBefore $_ $sprint.endDate)){
			$item.started="InSprint"
		}
		else {
			$item.started="NotStarted"
		}
		
		if (Check-DoneBefore $_ $sprint.startDate){
			$item.done="BeforeSprint"
		}
		elseif (Check-DoneAfter $_ $sprint.endDate){
			$item.done="AfterSprint"
		}
		elseif ((Check-DoneAfter $_ $sprint.startDate) -and (Check-DoneBefore $_ $sprint.endDate)){
			$item.done="InSprint"
		}
		else {
			$item.done="Undone"
		}
		$item.ever_assigned=Check-EverAssigned $_ $sprint
		$obj = New-Object PSObject -property $item
		$obj
	}
	
	if ([System.String]::IsNullOrEmpty($path)) {
		return $items
	}
	else {
		$items | select -Property id,title,type,state,revisions,sprint,assigned_date,assigned_by,ever_assigned,created_date,created_by,activated_date,activated_by,resolved_date,resolved_by,closed_date,closed_by,done_date,done_by,changed_date,change_by,planned,removed,started,done  | Export-Csv -Path $path -NoTypeInformation -Delimiter ';'
	}
}

Export-ModuleMember -Function Install-Az
Export-ModuleMember -Function Check-*
Export-ModuleMember -Function Get-*
Export-ModuleMember -Function New-*
