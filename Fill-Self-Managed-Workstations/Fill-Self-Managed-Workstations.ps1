#requires -version 2.0
###############################################################################
# Render SQL Reports using PowerShell
# This script is using the new Posh v 2.0 cmdlet New-WebServiceProxy
# FileName: RenderSQLReportFromPosh.v1.000.ps1
# Authors: Stefan Stranger (Microsoft)
# Help from: Jin Chen (Microsoft)
# Example of Rendering the OpsMgr Licenses Report from the Generic Report Library
#
# v1.000 – 15/05/2010 - stefstr - initial sstranger's release
# 10/29/2013 - Modified by Kirill 'kf' Nikolaev to satisfy SCCM needs
###############################################################################

#Log file path
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
$ScriptName = ($MyInvocation.MyCommand.Name).Substring(0,($MyInvocation.MyCommand.Name).Length-4)
$Log = Join-Path $ScriptPath "$ScriptName.txt"

Add-Content $Log (Get-Date)

#Clear-Content $Log -ErrorAction SilentlyContinue
$Error.Clear()
			
#Define Variables            
#Enter URI to asmx file on Report Server            
$URI  = 'https://reportserver.example.com//ReportServer//ReportExecution2005.asmx?wsdl'
#Enter Report Path
$ReportPath = '/ConfigMgr_CAS/Compliance and Settings Management/Workstations with non-default local administrators'
#Enter group name to fill with workstations
$GroupName = 'Self-Managed Workstations'

$format = 'csv'
$deviceinfo = ''
$extention = ''
$mimeType = ''
$Result = ''
$render = ''
$encoding = 'UTF-8'
$warnings = $null            
$streamIDs = $null            
$Reports = New-WebServiceProxy -Uri $URI -UseDefaultCredential -namespace 'ReportExecution2005'
            
            
$rsExec = new-object ReportExecution2005.ReportExecutionService            
$rsExec.Credentials = [System.Net.CredentialCache]::DefaultCredentials             

$execInfo = @($ReportPath, $null)             

#Load the selected report.            
$rsExec.GetType().GetMethod('LoadReport').Invoke($rsExec, $execInfo) | out-null              

#Report Parameters            
#Depending on the number of Parameters being used in the Report you need to add more Parameters.            
#Search the rdl file for the correct parameter names.
$param1 = new-object ReportExecution2005.ParameterValue
$param1.Name = 'CompOU'
$param1.Value = 'EXAMPLE.COM/WORKSTATIONS'

$param2 = new-object ReportExecution2005.ParameterValue
$param2.Name = 'CompOU'
$param2.Value = 'EXAMPLE.COM/WORKSTATIONS-OLD'

$parameters = [ReportExecution2005.ParameterValue[]] ($param1, $param2)             
$ExecParams = $rsExec.SetExecutionParameters($parameters, 'en-us');
$render = $rsExec.Render($format, $deviceInfo,[ref] $extention, [ref] $mimeType,[ref] $encoding, [ref] $warnings, [ref] $streamIDs)             
$Result = [text.encoding]::ascii.getString($render)             

$ComputerNames = @()
$SplittedResult = $Result.Split(“`n”)
for ($i = 1; $i -le $SplittedResult.Count-3) { #For unknown reason, last 3 rows of SSRS answer are blank, so we have to cut them and a title too.
	$ComputerNames += ($SplittedResult[$i]).Substring(0,($SplittedResult[$i]).Length-1) #Again, for unknown reason, there is an invisible line-feed symbol, we remove it.
	$i++
}
$ComputerObjects = @()
foreach ($Computer in $ComputerNames) {
	$ComputerObjects += Get-ADComputer $Computer
}
$CurrentMembers = @()
$CurrentMembers = Get-ADGroupMember -Identity $GroupName

$ToAdd = @()
$ToRemove = @()
if ($CurrentMembers) {
	if ($ComputerObjects) {
		$CompareResult = Compare-Object -ReferenceObject $ComputerObjects -DifferenceObject $CurrentMembers
		foreach ($Item in $CompareResult) {
			$DN = $Item.InputObject.DistinguishedName
			if ($Item.SideIndicator -eq '<=') {
				$ToAdd += $Item.InputObject
				Add-Content $Log "$DN - ToAdd"
			}
			elseif ($Item.SideIndicator -eq '=>') {
				$ToRemove += $Item.InputObject
				Add-Content $Log "$DN - ToRemove"
			}
		}
	}
	else {
		foreach ($Item in $CompareResult) {
			$DN = $Item.DistinguishedName
			$ToRemove += $Item.DistinguishedName
			Add-Content $Log "$DN - ToRemove"
		}
	}
}
else {
	foreach ($Item in $ComputerObjects) {
		$DN = $Item.DistinguishedName
		$ToAdd += $Item.DistinguishedName
		Add-Content $Log "$DN - ToAdd"
	}
}
if ($ToAdd) {
	try {
		Add-ADGroupMember -Identity $GroupName -Members $ToAdd -ErrorAction SilentlyContinue
	}
	catch {
		Add-Content $Log 'Cannot add'
		Add-Content $Log $Error[0]
	}
}
if ($ToRemove) {
	try {
		Remove-ADGroupMember -Identity $GroupName -Members $ToRemove -ErrorAction SilentlyContinue -Confirm:$false
		}
	catch {
		Add-Content $Log 'Cannot remove'
		Add-Content $Log $Error[0]
	}
}