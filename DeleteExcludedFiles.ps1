<#
.SYNOPSIS
Find all files excluded from a Visual Studio solution with options to delete.
.DESCRIPTION
Finds all excluded files in all projects in the provided Visual Studio solution with options to delete the files.
.PARAMETER Solution
The path to the .sln file
.PARAMETER VsVersion
The Visual Studio version (10, 11, 12) (Used to locate the tf.exe file)
.PARAMETER DeleteFromTfs
Mark files as pending deletion in TFS
.PARAMETER DeleteFromDisk
Delete the files directly from the disk
#>

[CmdletBinding()]
param(
    [Parameter(Position=0, Mandatory=$true)]
    [string]$Solution,
    [Parameter(Mandatory=$false)]
    [ValidateRange(10,12)] 
    [int] $VsVersion = 12,  
    [switch]$DeleteFromDisk,
    [switch]$DeleteFromTfs
)
$ErrorActionPreference = "Stop"
$tfPath = "${env:ProgramFiles(X86)}\Microsoft Visual Studio $VsVersion.0\Common7\IDE\TF.exe"
$solutionDir = Split-Path $Solution | % { (Resolve-Path $_).Path }

$projects = Select-String -Path $Solution -Pattern 'Project.*"(?<file>.*\.csproj)".*' `
	| % { $_.Matches[0].Groups[1].Value } `
	| % { Join-Path $solutionDir $_ }
	
$excluded = $projects | % {
	$projectDir = Split-Path $_

	$projectFiles = Select-String -Path $_ -Pattern '<(Compile|None|Content|EmbeddedResource) Include="(.*)".*' `
		| % { $_.Matches[0].Groups[2].Value } `
		| % { Join-Path $projectDir $_ }
		
	$diskFiles = Get-ChildItem -Path $projectDir -Recurse `
		| ? { !$_.PSIsContainer } `
		| % { $_.FullName } `
		| ? { $_ -notmatch "\\obj\\|\\bin\\|\\logs\\|\.user|\.*proj|App_Configuration\\|App_Data\\" }
		
	(compare-object $diskFiles $projectFiles -PassThru) | Where { $_.SideIndicator -eq '<=' }
} 

Write-Host "Found" $excluded.count "excluded files"

if ($DeleteFromTfs) 
{
	Write-Host "Marking excluded files as deleted in TFS..."
	$excluded | % {
		[Array]$arguments = @("delete", "`"$_`"")
		& "$tfPath" $arguments
	}
} 
elseif($DeleteFromDisk)
{
	Write-Host "Deleting excluded files from disk..."
	$excluded | % { Remove-Item -Path $_ -Force -Verbose}
}
else 
{
	Write-Host "Neither DeleteFromTfs or DeleteFromDisk was specified. Listing excluded files only..."
	$excluded
}