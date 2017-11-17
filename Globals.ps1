#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------


#Sample function that provides the location of the script
function Get-ScriptDirectory
{
<#
	.SYNOPSIS
		Get-ScriptDirectory returns the proper location of the script.

	.OUTPUTS
		System.String
	
	.NOTES
		Returns the correct path within a packaged executable.
#>
	[OutputType([string])]
	param ()
	if ($null -ne $hostinvocation)
	{
		Split-Path $hostinvocation.MyCommand.path
	}
	else
	{
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}

#Sample variable that provides the location of the script
[string]$ScriptDirectory = Get-ScriptDirectory

function Affiche_texte ($parameter1)
{
	[void][System.Windows.Forms.MessageBox]::Show($parameter1)
}
function Affiche_Erreur ($parameter1)
{
	$Title = "Erreur..."
	$Button = [Windows.Forms.MessageBoxButtons]::OK
	$Icon = [Windows.Forms.MessageBoxIcon]::Information
	
	[void][System.Windows.Forms.MessageBox]::Show($parameter1, $Title, $Button, $Icon)
}

<#
	.SYNOPSIS
		Get the latest Cumulative update for Windows
	
	.DESCRIPTION
		This script will return the list of Cumulative updates for Windows 10 and Windows Server 2016 from the Microsoft Update Catalog.
	
	.PARAMETER StartKB
		JSON source for the update KB articles.
	
	.PARAMETER Build
		Windows 10 Build Number used to filter avaible Downloads
		
		10240 - Windows 10 Version 1507
		10586 - Windows 10 Version 1511
		14393 - Windows 10 Version 1607 and Windows Server 2016
		15063 - Windows 10 Version 1703
	
	.PARAMETER Filter
		Specify a specific search filter to change the target update behaviour. The default will only Cumulative updates for x86 and x64.
		
		If Mulitple Filters are specified, only string that match *ALL* filters will be selected.
		
		Cumulative - Download Cumulative updates.
		Delta - Download Delta updates.
		x86 - Download x86
		x64 - Download x64
	
	.PARAMETER Force
		A description of the Force parameter.
	
	.EXAMPLE
		Get the latest Cumulative Update for Windows 10 x64
		
		.\Get-LatestUpdate.ps1
	
	.EXAMPLE
		Get the latest Cumulative Update for Windows 10 x86
		
		.\Get-LatestUpdate.ps1 -Filter 'Cumulative','x86'
	
	.EXAMPLE
		Get the latest Cumulative Update for Windows Server 2016
		
		.\Get-LatestUpdate.ps1 -Filter 'Cumulative','x64' -Build 1607
	
	.EXAMPLE
		Get the latest Cumulative Updates for Windows 10 (both x86 and x64) and download to the %TEMP% directory.
		
		.\Get-LatestUpdate.ps1 | Start-BitsTransfer -Destination $env:Temp
	
	.NOTES
		Copyright Keith Garner (KeithGa@DeploymentLive.com), All rights reserved.
		
		
#>
function Get-LatestUpdate
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $false,
				   HelpMessage = 'JSON source for the update KB articles.')]
		[string]$StartKB = 'https://support.microsoft.com/app/content/api/content/asset/en-us/4000816',
		[Parameter(Mandatory = $false,
				   HelpMessage = 'Windows build number.')]
		[ValidateSet('16299','15063', '14393', '10586', '10240')]
		[string]$Build = '15063',
		[Parameter(Mandatory = $false,
				   HelpMessage = 'Windows update Catalog Search Filter.')]
		[ValidateSet('x64', 'x86', 'Cumulative', 'Delta')]
		[string[]]$Filter = @("Cumulative")
	)
	
	#region Support Routine
	
	Function Select-LatestUpdate
	{
		[CmdletBinding(SupportsShouldProcess = $True)]
		Param (
			[Parameter(Mandatory = $true, ValueFromPipeline = $true)]
			$Updates
		)
		Begin
		{
			$MaxObject = $null
			$MaxValue = [version]::new("0.0")
		}
		Process
		{
			ForEach ($Update in $updates)
			{
				Select-String -InputObject $Update -AllMatches -Pattern "(\d+\.)?(\d+\.)?(\d+\.)?(\*|\d+)" |
				ForEach-Object { $_.matches.value } |
				ForEach-Object { $_ -as [version] } |
				ForEach-Object {
					if ($_ -gt $MaxValue) { $MaxObject = $Update; $MaxValue = $_ }
				}
			}
		}
		End
		{
			$MaxObject | Write-Output
		}
	}
	
	#endregion
	
	#region Find the KB Article Number
	
	Write-Verbose "Downloading $StartKB to retrieve the list of updates."
	$kbID = Invoke-WebRequest -Uri $StartKB |
	Select-Object -ExpandProperty Content |
	ConvertFrom-Json |
	Select-Object -ExpandProperty Links |
	Where-Object level -eq 2 |
	Where-Object text -match $Build |
	Select-LatestUpdate |
	Select-Object -First 1
	
	#endregion
	
	#region get the download link from Windows Update
	
	Write-Verbose "Found ID: KB$($kbID.articleID)"
	$kbObj = Invoke-WebRequest -Uri "http://www.catalog.update.microsoft.com/Search.aspx?q=KB$($KBID.articleID)"
	
	$Available_KBIDs = $kbObj.InputFields |
	Where-Object { $_.type -eq 'Button' -and $_.Value -eq 'Download' } |
	Select-Object -ExpandProperty ID
	
	$Available_KBIDs | out-string | write-verbose
	
	$kbGUIDs = $kbObj.Links |
	Where-Object ID -match '_link' |
	Where-Object { $_.OuterHTML -match ("(?=.*" + ($Filter -join ")(?=.*") + ")") } |
	ForEach-Object { $_.id.replace('_link', '') } |
	Where-Object { $_ -in $Available_KBIDs }
	
	foreach ($kbGUID in $kbGUIDs)
	{
		Write-Verbose "`t`tDownload $kbGUID"
		$Post = @{ size = 0; updateID = $kbGUID; uidInfo = $kbGUID } | ConvertTo-Json -Compress
		$PostBody = @{ updateIDs = "[$Post]" }
		Invoke-WebRequest -Uri 'http://www.catalog.update.microsoft.com/DownloadDialog.aspx' -Method Post -Body $postBody |
		Select-Object -ExpandProperty Content |
		Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)" |
		Select-Object -Unique |
		ForEach-Object { [PSCustomObject] @{ Source = $_.matches.value } } # Output for BITS
	}
}

function Get-OsCulture
	{
		[CmdletBinding()]
		param ()
		
		$oscode = Get-WmiObject Win32_OperatingSystem -ComputerName localhost -ErrorAction continue | foreach { $_.oslanguage }
		$Culture = switch ($oscode) `
		{
			1033 { "English" };
			1036 { "French" };
			default { "Unknown" }
		}
		
		switch ($Culture)
		{
			'French' {
				
				$global:OsCulture = "French"
				$global:OsCultureValue = "FR-FR"
				return $global:OsCulture, $global:OsCultureValue
				
				
			}
			{ 'English' } {
				
				$global:OsCulture = "English"
				$global:OsCultureValue = "EN-US"
				return $global:OsCulture, $global:OsCultureValue
			}
			
			Default { }
		}
	}
function Get-Ram
{
		[CmdletBinding()]
		param ()
		
		[string]$computername = "."
		$memorymeasure = Get-WMIObject Win32_PhysicalMemory -ComputerName $computername | Measure-Object -Property Capacity -Sum
		
		#  Format and Print
		$Global:InfoRam = "{0} GB" -f $($memorymeasure.sum/1024/1024/1024)
		
	}
function Found-ADK
	{
		[CmdletBinding()]
		param
		(
			[Parameter(Mandatory = $true)]
			$OsCulture = 'FR-FR'
		)
		switch ($OsCulture)
		{
			'EN-US'
			{
				$ADK = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Windows Assessment and Deployment Kit - Windows 10" }
				if ($ADK -eq $null)
				{
					$global:ADKpresent = 'Not Found'
					return $global:ADKpresent
				}
				else
				{
					$global:ADKpresent = 'Present'
					$global:ADKv = $ADK.DisplayVersion
					return $global:ADKpresent, $global:ADKv
				}
			}
			'FR-FR'
			{
				$ADK = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "Kit* Windows 10" }
				if ($ADK -eq $null)
				{
					$global:ADKpresent = 'Not Found'
					$global:ADKv = $ADK.DisplayVersion
					return $global:ADKpresent, $global:ADKv
				}
				else
				{
					$global:ADKpresent = 'Present'
					$global:ADKv = $ADK.DisplayVersion
					return $global:ADKpresent, $global:ADKv
				}
			}
			
			Default { }
		}
		
	}
	
function Found-MDT
	{
		[CmdletBinding()]
		param
		(
			[Parameter(Mandatory = $true)]
			$OsCulture = 'FR-FR'
		)
		switch ($OsCulture)
		{
			'EN-US'
			{
				$MDT = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Microsoft Deployment Toolkit (6.3.8443.1000)" }
				if ($MDT -eq $null)
				{
					$global:MDTpresent = 'Not Found'
					return $global:MDTpresent
				}
				else
				{
					$global:PathMDT = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Deployment 4" -Name Install_Dir).Install_Dir
					$global:MDTpresent = 'Present'
					$global:MDTv = $MDT.DisplayVersion
					return $global:MDTpresent, $global:MDTv, $global:PathMDT
				}
			}
			'FR-FR'
			{
				$MDT = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Microsoft Deployment Toolkit (6.3.8443.1000)" }
				if ($MDT -eq $null)
				{
					$global:MDTpresent = 'Not Found'
					return $global:MDTpresent
				}
				else
				{
					$global:PathMDT = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Deployment 4" -Name Install_Dir).Install_Dir
					$global:MDTpresent = 'Present'
					$global:MDTv = $MDT.DisplayVersion
					return $global:MDTpresent, $global:MDTv, $global:PathMDT
				}
			}
			
			Default { }
		}
		
	}
	
function Write-log
	{
		param
		(
			[Parameter(Mandatory = $true)]
			[String]$message
		)
		
		$timeStamp = Get-Date -Format "[dd/MM/yy - %H:mm:ss]"
		$VerboseLogFile = "$ScriptDirectory\CreateReferenceImageMDT.log"
		$global:Log = Write-Output "$timestamp $message"
		$Logmessage = "$timestamp $message"
		$logMessage | Out-File -Append -LiteralPath $verboseLogFile
	}
	
	
	#region Control Helper Functions
	function Update-ListBox
	{
<#
	.SYNOPSIS
		This functions helps you load items into a ListBox or CheckedListBox.
	
	.DESCRIPTION
		Use this function to dynamically load items into the ListBox control.
	
	.PARAMETER ListBox
		The ListBox control you want to add items to.
	
	.PARAMETER Items
		The object or objects you wish to load into the ListBox's Items collection.
	
	.PARAMETER DisplayMember
		Indicates the property to display for the items in this control.
	
	.PARAMETER Append
		Adds the item(s) to the ListBox without clearing the Items collection.
	
	.EXAMPLE
		Update-ListBox $ListBox1 "Red", "White", "Blue"
	
	.EXAMPLE
		Update-ListBox $listBox1 "Red" -Append
		Update-ListBox $listBox1 "White" -Append
		Update-ListBox $listBox1 "Blue" -Append
	
	.EXAMPLE
		Update-ListBox $listBox1 (Get-Process) "ProcessName"
	
	.NOTES
		Additional information about the function.
#>
		
		param
		(
			[Parameter(Mandatory = $true)]
			[ValidateNotNull()]
			[System.Windows.Forms.ListBox]$ListBox,
			[Parameter(Mandatory = $true)]
			[ValidateNotNull()]
			$Items,
			[Parameter(Mandatory = $false)]
			[string]$DisplayMember,
			[switch]$Append
		)
		
		if (-not $Append)
		{
			$listBox.Items.Clear()
		}
		
		if ($Items -is [System.Windows.Forms.ListBox+ObjectCollection] -or $Items -is [System.Collections.ICollection])
		{
			$listBox.Items.AddRange($Items)
		}
		elseif ($Items -is [System.Collections.IEnumerable])
		{
			$listBox.BeginUpdate()
			foreach ($obj in $Items)
			{
				$listBox.Items.Add($obj)
				$listbox.SelectedIndex = $listbox.Items.Count - 1
				$listbox.ScrollIntoView($listbox.SelectedItem)
				
			}
			$listBox.EndUpdate()
		}
		else
		{
			$listBox.Items.Add($Items)
		}
		
		$listBox.DisplayMember = $DisplayMember
	}
	
	
	function Update-ComboBox
	{
<#
	.SYNOPSIS
		This functions helps you load items into a ComboBox.
	
	.DESCRIPTION
		Use this function to dynamically load items into the ComboBox control.
	
	.PARAMETER ComboBox
		The ComboBox control you want to add items to.
	
	.PARAMETER Items
		The object or objects you wish to load into the ComboBox's Items collection.
	
	.PARAMETER DisplayMember
		Indicates the property to display for the items in this control.
	
	.PARAMETER Append
		Adds the item(s) to the ComboBox without clearing the Items collection.
	
	.EXAMPLE
		Update-ComboBox $combobox1 "Red", "White", "Blue"
	
	.EXAMPLE
		Update-ComboBox $combobox1 "Red" -Append
		Update-ComboBox $combobox1 "White" -Append
		Update-ComboBox $combobox1 "Blue" -Append
	
	.EXAMPLE
		Update-ComboBox $combobox1 (Get-Process) "ProcessName"
	
	.NOTES
		Additional information about the function.
#>
		
		param
		(
			[Parameter(Mandatory = $true)]
			[ValidateNotNull()]
			[System.Windows.Forms.ComboBox]$ComboBox,
			[Parameter(Mandatory = $true)]
			[ValidateNotNull()]
			$Items,
			[Parameter(Mandatory = $false)]
			[string]$DisplayMember,
			[switch]$Append
		)
		
		if (-not $Append)
		{
			$ComboBox.Items.Clear()
		}
		
		if ($Items -is [Object[]])
		{
			$ComboBox.Items.AddRange($Items)
		}
		elseif ($Items -is [System.Collections.IEnumerable])
		{
			$ComboBox.BeginUpdate()
			foreach ($obj in $Items)
			{
				$ComboBox.Items.Add($obj)
			}
			$ComboBox.EndUpdate()
		}
		else
		{
			$ComboBox.Items.Add($Items)
		}
		
		$ComboBox.DisplayMember = $DisplayMember
	}
	#endregion
function Get-DiskFree
	{
		[CmdletBinding()]
		param ()
		
		$NeededFreeSpace = 50 #GigaBytes
		$Disk = Get-wmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
		$FreeSpace = [MATH]::ROUND($disk.FreeSpace /1GB)
		$global:Diskavailable = "$FreeSpace Go"
	}
function New-RefWimMDT
	{
		[CmdletBinding(DefaultParameterSetName = '1',
					   SupportsShouldProcess = $true)]
		[OutputType([psobject])]
		param
		(
			[Parameter(Mandatory = $true)]
			$SourceISO,
			[Parameter(Mandatory = $true)]
			$Update,
			[Parameter(Mandatory = $false)]
			$MountFolder,
			[Parameter(Mandatory = $true)]
			[ValidateSet('REFWS2K16S-001.wim', 'REFWS2K16SC-002.wim', 'REFWS2K16D-003.wim', 'REFWS2K16DC-004.wim', 'REFWS2K16Full.wim', 'REFW10PRO.wim', 'REFW10EDU.wim', 'REFW10ENT.wim')]
			$RefImage = 'REFWS2K16S-001.wim',
			[Parameter(Mandatory = $false)]
			$Path,
			[ValidateSet('YES', 'NO')]
			$ISO = 'NO',
			[ValidateSet('YES', 'NO')]
			$NanoUpdate,
			[ValidateSet('YES', 'NO')]
			$RemoveNano
		)
		
		$StartTime = Get-Date
		$MountFolder = $MountFolder +"\"
		if (!(Test-Path -path $Path)) { New-Item $Path -Type Directory | Out-Null }
		
		# Verify that the ISO and CU files existnote
		#if (!(Test-Path -path $SourceISO)) { Write-Warning "Could not find Windows Server 2016 ISO file. Aborting..."; Write-log -message "Could not find Windows Server 2016 ISO file. Aborting..."; Break }
		#if (!(Test-Path -path $Update)) { Write-Warning "Cumulative Update for Windows Server 2016. Aborting..."; Write-log -message "Cumulative Update for Windows Server 2016. Aborting..."; Break }
		
		# Mount the Windows Server 2016 ISO
		Mount-DiskImage -ImagePath $SourceISO
		[System.Windows.Forms.Application]::DoEvents()
		Write-log -message "Mount Windows Operating ISO "
		Update-ListBox -Items $Log -Append -ListBox $listbox1
		$ISOImage = Get-DiskImage -ImagePath $SourceISO | Get-Volume
		$ISODrive = [string]$ISOImage.DriveLetter + ":"
		$FPath = $Path + "\" + $RefImage
		if (!(Test-Path -path $Path)) { New-Item -path $Path -ItemType Directory | Out-Null }
		
		switch ($RefImage)
		{
			REFWS2K16S-001.wim {
				
				Write-log -message "Extract the Windows Server 2016 Standard index to a new WIM"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows Server 2016 SERVERSTANDARD" -DestinationImagePath $FPath
				Write-log -message "Create Mount Folder"
				[System.Windows.Forms.Application]::DoEvents()
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				if (!(Test-Path -path $MountFolder)) { New-Item -path $MountFolder -ItemType Directory | Out-Null }
				Write-log -message "Mount $RefImage Image..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
				
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Disable Protocol SMBv1"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				
				LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
				[System.Windows.Forms.Application]::DoEvents()
				Wait-Process -Name Dism
				[System.Windows.Forms.Application]::DoEvents()
				
				Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				Add-WindowsPackage -PackagePath $Update -Path $MountFolder
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Add .NET Framework 3.5.1 to the Windows Server 2016 Standard image "
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Dismount the Windows Server 2016 Standard image"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				DisMount-WindowsImage -Path $MountFolder -Save
				[System.Windows.Forms.Application]::DoEvents()
				$Path2 = $Path +"\"+ "REFWS2K16S-001Inter.wim"
				Write-log -message "Export WimFile to reduce the size for Windows Server 2016 Standard "
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Export-WindowsImage -SourceImagePath $FPath -CompressionType maximum -SourceName "Windows Server 2016 SERVERSTANDARD" -DestinationImagePath $Path2
				[System.Windows.Forms.Application]::DoEvents()
				Remove-Item -Path $FPath -Force
				[System.Windows.Forms.Application]::DoEvents()
				Rename-Item -Path "$Path2" -NewName "$RefImage"
				Write-log -message "Windows Server 2016 Standard Image is finished"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				
				
			}
			REFWS2K16SC-002.wim {
				
				Write-log -message "Extract the Windows Server 2016 Standard Core index to a new WIM"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows Server 2016 SERVERSTANDARDCORE" -DestinationImagePath $FPath
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Create Mount Folder"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				if (!(Test-Path -path $MountFolder)) { New-Item -path $MountFolder -ItemType Directory | Out-Null }
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Mount $RefImage Image..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
				
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Disable Protocol SMBv1"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				
				LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
				[System.Windows.Forms.Application]::DoEvents()
				Wait-Process -Name Dism
				[System.Windows.Forms.Application]::DoEvents()
				
				Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				Add-WindowsPackage -PackagePath $Update -Path $MountFolder
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Add .NET Framework 3.5.1 to the Windows Server 2016 Standard image "
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Dismount the Windows Server 2016 Standard image"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				DisMount-WindowsImage -Path $MountFolder -Save
				[System.Windows.Forms.Application]::DoEvents()
			    $Path2 = $Path + "\" + "REFWS2K16SC-002Inter.wim"
				Write-log -message "Export WimFile to reduce the size for Windows Server 2016 Standard Core"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Export-WindowsImage -SourceImagePath $FPath -CompressionType maximum -SourceName "Windows Server 2016 SERVERSTANDARDCORE" -DestinationImagePath $Path2
				[System.Windows.Forms.Application]::DoEvents()
				Remove-Item -Path $FPath -Force
				[System.Windows.Forms.Application]::DoEvents()
				Rename-Item -Path "$Path2" -NewName "$RefImage"
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Windows Server 2016 Standard Core Image is finished"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
			}
			REFWS2K16D-003.wim {
				
				Write-log -message "Extract the Windows Server 2016 Datacenter index to a new WIM"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows Server 2016 SERVERDATACENTER" -DestinationImagePath $FPath
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Create Mount Folder"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				if (!(Test-Path -path $MountFolder)) { New-Item -path $MountFolder -ItemType Directory | Out-Null }
				Write-log -message "Mount $RefImage Image..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
				
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Disable Protocol SMBv1"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				
				LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
				[System.Windows.Forms.Application]::DoEvents()
				Wait-Process -Name Dism
				[System.Windows.Forms.Application]::DoEvents()
				
				Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				Add-WindowsPackage -PackagePath $Update -Path $MountFolder
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Add .NET Framework 3.5.1 to the Windows Server 2016 Datacenter image "
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Dismount the Windows Server 2016 Standard image"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				DisMount-WindowsImage -Path $MountFolder -Save
				[System.Windows.Forms.Application]::DoEvents()
			    $Path2 = $Path + "\"+ "REFWS2K16D-003Inter.wim"
				Write-log -message "Export WimFile to reduce the size for Windows Server 2016 Datacenter "
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Export-WindowsImage -SourceImagePath $FPath -CompressionType maximum -SourceName "Windows Server 2016 SERVERDATACENTER" -DestinationImagePath $Path2
				[System.Windows.Forms.Application]::DoEvents()
				Remove-Item -Path $FPath -Force
				[System.Windows.Forms.Application]::DoEvents()
				Rename-Item -Path "$Path2" -NewName "$RefImage"
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Windows Server 2016 Datacenter Image is finished"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
			}
			REFWS2K16DC-004.wim {
				
				Write-log -message "Extract the Windows Server 2016 Datacenter Core index to a new WIM"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows Server 2016 SERVERDATACENTERCORE" -DestinationImagePath $FPath
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Create Mount Folder"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				if (!(Test-Path -path $MountFolder)) { New-Item -path $MountFolder -ItemType Directory | Out-Null }
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Mount $RefImage Image..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
				
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Disable Protocol SMBv1"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				
				LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
				[System.Windows.Forms.Application]::DoEvents()
				Wait-Process -Name Dism
				[System.Windows.Forms.Application]::DoEvents()
				
				Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				Add-WindowsPackage -PackagePath $Update -Path $MountFolder
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Add .NET Framework 3.5.1 to the Windows Server 2016 Datacenter Core image "
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Dismount the Windows Server 2016 Datacenter Core image"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				DisMount-WindowsImage -Path $MountFolder -Save
				[System.Windows.Forms.Application]::DoEvents()
			    $Path2 = $Path + "\"+ "REFWS2K16DC-004Inter.wim"
				Write-log -message "Export WimFile to reduce the size for Windows Server 2016 Datacenter Core"
				[System.Windows.Forms.Application]::DoEvents()
				Export-WindowsImage -SourceImagePath $FPath -CompressionType maximum -SourceName "Windows Server 2016 SERVERDATACENTERCORE" -DestinationImagePath $Path2
				[System.Windows.Forms.Application]::DoEvents()
				Remove-Item -Path $FPath -Force
				[System.Windows.Forms.Application]::DoEvents()
				Rename-Item -Path "$Path2" -NewName "$RefImage"
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Windows Server 2016 Datacenter Core Image is finished"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
			}
			REFWS2K16Full.wim{
				if ($ISO -eq 'NO')
				{
					Write-log -message "========================================"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					Write-log -message "Full DVD Update Proceed for all index..."
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					Write-log -message "========================================"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					Write-log -message "Edition Server 2016 Standard"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					Write-log -message "Extract the Windows Server 2016 Standard index to a new WIM"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					$FPath = $Path +"\"+ "REFWS2K16S-001.wim"
					[System.Windows.Forms.Application]::DoEvents()
					Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows Server 2016 SERVERSTANDARD" -DestinationImagePath $FPath
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Create Mount Folder"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					if (!(Test-Path -path $MountFolder)) { New-Item -path $MountFolder -ItemType Directory | Out-Null }
					Write-log -message "Mount REFWS2K16S-001.wim Image..."
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					[System.Windows.Forms.Application]::DoEvents()
					Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
					
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Disable Protocol SMBv1"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					
					LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
					[System.Windows.Forms.Application]::DoEvents()
					Wait-Process -Name Dism
					[System.Windows.Forms.Application]::DoEvents()
					
					Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					Add-WindowsPackage -PackagePath $Update -Path $MountFolder
					[System.Windows.Forms.Application]::DoEvents()
					
					Write-log -message "Add .NET Framework 3.5.1 to the Windows Server 2016 Standard image "
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					[System.Windows.Forms.Application]::DoEvents()
					Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
					[System.Windows.Forms.Application]::DoEvents()
					
					Write-log -message "Dismount the Windows Server 2016 Standard image"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					[System.Windows.Forms.Application]::DoEvents()
					DisMount-WindowsImage -Path $MountFolder -Save
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Edition Server 2016 Standard Core"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					Write-log -message "Extract the Windows Server 2016 Standard Core index to a new WIM"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					$FPath = $Path +"\"+ "REFWS2K16SC-002.wim"
					[System.Windows.Forms.Application]::DoEvents()
					Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows Server 2016 SERVERSTANDARDCORE" -DestinationImagePath $FPath
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Mount REFWS2K16SC-002.wim Image..."
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					[System.Windows.Forms.Application]::DoEvents()
					Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
					
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Disable Protocol SMBv1"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					
					LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
					[System.Windows.Forms.Application]::DoEvents()
					Wait-Process -Name Dism
					[System.Windows.Forms.Application]::DoEvents()
					
					Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					Add-WindowsPackage -PackagePath $Update -Path $MountFolder
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Add .NET Framework 3.5.1 to the Windows Server 2016 Standard Core image "
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					[System.Windows.Forms.Application]::DoEvents()
					Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Dismount the Windows Server 2016 Standard Core image"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					[System.Windows.Forms.Application]::DoEvents()
					DisMount-WindowsImage -Path $MountFolder -Save
					[System.Windows.Forms.Application]::DoEvents()
					
					Write-log -message "Edition Server 2016 Datacenter "
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					Write-log -message "Extract the Windows Server 2016 Datacenter index to a new WIM"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					$FPath = $Path +"\"+ "REFWS2K16D-003.wim"
					[System.Windows.Forms.Application]::DoEvents()
					Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows Server 2016 SERVERDATACENTER" -DestinationImagePath $FPath
					[System.Windows.Forms.Application]::DoEvents()
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					Write-log -message "Mount REFWS2K16D-003.wim Image..."
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					[System.Windows.Forms.Application]::DoEvents()
					Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
					
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Disable Protocol SMBv1"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					
					LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
					[System.Windows.Forms.Application]::DoEvents()
					Wait-Process -Name Dism
					[System.Windows.Forms.Application]::DoEvents()
					
					Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					Add-WindowsPackage -PackagePath $Update -Path $MountFolder
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Add .NET Framework 3.5.1 to the Windows Server 2016 Datacenter image "
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					[System.Windows.Forms.Application]::DoEvents()
					Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Dismount the Windows Server 2016 Datacenter image"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					[System.Windows.Forms.Application]::DoEvents()
					DisMount-WindowsImage -Path $MountFolder -Save
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Edition Server 2016 Datacenter Core"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					Write-log -message "Extract the Windows Server 2016 Datacenter Core index to a new WIM"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					$FPath = $Path +"\"+ "REFWS2K16DC-004.wim"
					[System.Windows.Forms.Application]::DoEvents()
					Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows Server 2016 SERVERDATACENTERCORE" -DestinationImagePath $FPath
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Mount REFWS2K16DC-004.wim Image..."
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					[System.Windows.Forms.Application]::DoEvents()
					Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
					
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Disable Protocol SMBv1"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					
					LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
					[System.Windows.Forms.Application]::DoEvents()
					Wait-Process -Name Dism
					[System.Windows.Forms.Application]::DoEvents()
					
					Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					Add-WindowsPackage -PackagePath $Update -Path $MountFolder
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Add .NET Framework 3.5.1 to the Windows Server 2016 Datacenter Core image "
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					[System.Windows.Forms.Application]::DoEvents()
					Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Dismount the Windows Server 2016 Datacenter Core image"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					[System.Windows.Forms.Application]::DoEvents()
					DisMount-WindowsImage -Path $MountFolder -Save
					[System.Windows.Forms.Application]::DoEvents()
					
					Write-log -message "All Editions are update. Merged all to Install.wim and optimize the size"
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					$Files = ('REFWS2K16S-001.wim', 'REFWS2K16SC-002.wim', 'REFWS2K16D-003.wim', 'REFWS2K16DC-004.wim')
					foreach ($item in $Files)
					{
						$Path3 = $Path+"\" + $item
						Write-log -message "export $item to install.wim"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						$SourceName = switch ($item)
						{
							REFWS2K16S-001.wim {
								"Windows Server 2016 SERVERSTANDARD"
							}
							REFWS2K16SC-002.wim {
								"Windows Server 2016 SERVERSTANDARDCORE"
							}
							REFWS2K16D-003.wim {
								"Windows Server 2016 SERVERDATACENTER"
							}
							REFWS2K16DC-004.wim{
								"Windows Server 2016 SERVERDATACENTERCORE"
							}
							default
							{
								#<code>
							}
						}
						$Path2 = $Path +"\"+ "Install.wim"
						[System.Windows.Forms.Application]::DoEvents()
						Export-WindowsImage -SourceImagePath $Path3 -CompressionType maximum -SourceName $SourceName -DestinationImagePath $Path2
						[System.Windows.Forms.Application]::DoEvents()
						
					}
					
				}
				else
				{
					
					Write-log -message "===================================================="
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					Write-log -message "Full DVD Update with ISO Proceed for all index..."
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					Write-log -message "===================================================="
					Update-ListBox -Items $Log -Append -ListBox $listbox1
					
					if ($global:ADKpresent -eq 'Present')
					{
						##############################################################################
						$Here = Get-Location
						$ADK_Path = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit"
						$WinPE_ADK_Path = $ADK_Path + "\Windows Preinstallation Environment"
						$WinPE_OCs_Path = $WinPE_ADK_Path + "\$WinPEArchitecture\WinPE_OCs"
						$DISM_Path = $ADK_Path + "\Deployment Tools" + "\amd64\Oscdimg"
						$ADK_Path_Media = $ADK_Path + $WinPE_ADK_Path + "\$WinPEArchitecture\Media"
						$NewISoName = Split-Path $SourceISO -Leaf
						
						Write-log -message "Edition Server 2016 Standard"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						Write-log -message "Extract the Windows Server 2016 Standard index to a new WIM"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						$FPath = $Path +"\"+ "REFWS2K16S-001.wim"
						[System.Windows.Forms.Application]::DoEvents()
						Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows Server 2016 SERVERSTANDARD" -DestinationImagePath $FPath
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Create Mount Folder"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						if (!(Test-Path -path $MountFolder)) { New-Item -path $MountFolder -ItemType Directory | Out-Null }
						Write-log -message "Mount REFWS2K16S-001.wim Image..."
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						[System.Windows.Forms.Application]::DoEvents()
						Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
						
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Disable Protocol SMBv1"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						
						LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
						[System.Windows.Forms.Application]::DoEvents()
						Wait-Process -Name Dism
						[System.Windows.Forms.Application]::DoEvents()
						
						Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						Add-WindowsPackage -PackagePath $Update -Path $MountFolder
						[System.Windows.Forms.Application]::DoEvents()
						
						Write-log -message "Add .NET Framework 3.5.1 to the Windows Server 2016 Standard image "
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						[System.Windows.Forms.Application]::DoEvents()
						Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Dismount the Windows Server 2016 Standard image"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						[System.Windows.Forms.Application]::DoEvents()
						DisMount-WindowsImage -Path $MountFolder -Save
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Edition Server 2016 Standard Core"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						Write-log -message "Extract the Windows Server 2016 Standard Core index to a new WIM"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						$FPath = $Path +"\"+ "REFWS2K16SC-002.wim"
						[System.Windows.Forms.Application]::DoEvents()
						Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows Server 2016 SERVERSTANDARDCORE" -DestinationImagePath $FPath
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Mount REFWS2K16SC-002.wim Image..."
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						[System.Windows.Forms.Application]::DoEvents()
						Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
						
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Disable Protocol SMBv1"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						
						LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
						[System.Windows.Forms.Application]::DoEvents()
						Wait-Process -Name Dism
						[System.Windows.Forms.Application]::DoEvents()
						
						Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						Add-WindowsPackage -PackagePath $Update -Path $MountFolder
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Add .NET Framework 3.5.1 to the Windows Server 2016 Standard Core image "
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						[System.Windows.Forms.Application]::DoEvents()
						Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Dismount the Windows Server 2016 Standard Core image"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						[System.Windows.Forms.Application]::DoEvents()
						DisMount-WindowsImage -Path $MountFolder -Save
						[System.Windows.Forms.Application]::DoEvents()
						
						Write-log -message "Edition Server 2016 Datacenter "
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						Write-log -message "Extract the Windows Server 2016 Datacenter index to a new WIM"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						$FPath = $Path +"\"+ "REFWS2K16D-003.wim"
						[System.Windows.Forms.Application]::DoEvents()
						Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows Server 2016 SERVERDATACENTER" -DestinationImagePath $FPath
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Mount REFWS2K16D-003.wim Image..."
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
						
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Disable Protocol SMBv1"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						
						LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
						[System.Windows.Forms.Application]::DoEvents()
						Wait-Process -Name Dism
						[System.Windows.Forms.Application]::DoEvents()
						
						Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						Add-WindowsPackage -PackagePath $Update -Path $MountFolder
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Add .NET Framework 3.5.1 to the Windows Server 2016 Datacenter image "
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						[System.Windows.Forms.Application]::DoEvents()
						Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Dismount the Windows Server 2016 Datacenter image"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						[System.Windows.Forms.Application]::DoEvents()
						DisMount-WindowsImage -Path $MountFolder -Save
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Edition Server 2016 Datacenter Core"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						Write-log -message "Extract the Windows Server 2016 Datacenter Core index to a new WIM"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						$FPath = $Path +"\"+ "REFWS2K16DC-004.wim"
						[System.Windows.Forms.Application]::DoEvents()
						Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows Server 2016 SERVERDATACENTERCORE" -DestinationImagePath $FPath
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Mount REFWS2K16DC-004.wim Image..."
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						[System.Windows.Forms.Application]::DoEvents()
						Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
						
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Disable Protocol SMBv1"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						
						LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
						[System.Windows.Forms.Application]::DoEvents()
						Wait-Process -Name Dism
						[System.Windows.Forms.Application]::DoEvents()
						
						Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						Add-WindowsPackage -PackagePath $Update -Path $MountFolder
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Add .NET Framework 3.5.1 to the Windows Server 2016 Datacenter Core image "
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						[System.Windows.Forms.Application]::DoEvents()
						Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Dismount the Windows Server 2016 Datacenter Core image"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						[System.Windows.Forms.Application]::DoEvents()
						DisMount-WindowsImage -Path $MountFolder -Save
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "All Editions are update. Merged all to Install.wim and optimize the size"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						[System.Windows.Forms.Application]::DoEvents()
						$Files = ('REFWS2K16S-001.wim', 'REFWS2K16SC-002.wim', 'REFWS2K16D-003.wim', 'REFWS2K16DC-004.wim')
						foreach ($item in $Files)
						{
							$Path3 = $Path + "\" + $item
							Write-log -message "export $item to install.wim"
							Update-ListBox -Items $Log -Append -ListBox $listbox1
							
							$SourceName = switch ($item)
							{
								REFWS2K16S-001.wim {
									"Windows Server 2016 SERVERSTANDARD"
								}
								REFWS2K16SC-002.wim {
									"Windows Server 2016 SERVERSTANDARDCORE"
								}
								REFWS2K16D-003.wim {
									"Windows Server 2016 SERVERDATACENTER"
								}
								REFWS2K16DC-004.wim{
									"Windows Server 2016 SERVERDATACENTERCORE"
								}
								default
								{
									#<code>
								}
							}
							$Path2 = $Path +"\"+ "Install.wim"
							[System.Windows.Forms.Application]::DoEvents()
							Export-WindowsImage -SourceImagePath $Path3 -CompressionType maximum -SourceName $SourceName -DestinationImagePath $Path2
							[System.Windows.Forms.Application]::DoEvents()
							
						}
						
						
						
						New-Item -path "$Path`ISO" -ItemType Directory | Out-Null
						$ISOFinal = $Path+"\" + "ISO"
						Write-log -message "Create ISO Structure et Copying files"
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						[System.Windows.Forms.Application]::DoEvents()
						Copy-Item "$ISODrive\*"  "$ISOFinal\" -Recurse -Exclude "install.wim"
						[System.Windows.Forms.Application]::DoEvents()
						Move-Item  $Path2 -Destination "$Path`ISO\sources\"
						[System.Windows.Forms.Application]::DoEvents()
						Write-log 'Create ISO with full update wim'
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						$VolumeName = (Get-WmiObject -Class Win32_CDROMDrive | Where-Object { $_.Caption -eq 'DVD-ROM virtuel Microsoft' }).volumeName
						#create Image ISO for UEFI and Bios
						
						Set-Location $DISM_Path
						$timeStampFile = Get-Date -Format "MM/dd/yyyy,%H:mm:ss" -Hour 23
						$BOOTDATA = "2#p0,e,b" + "$ISOFinal\boot\etfsboot.com" + "#pEF,e,b" + "$ISOFinal\efi\microsoft\boot\efisys_noprompt.bin"
						$ArgumentList = "-m" + " -o" + " -u2" + " -udfver102 " + " -l$VolumeName " + " -t$timeStampFile " + "-bootdata:$BOOTDATA", " $ISOFinal" + " $Path$NewISoName"
						start-Process ".\oscdimg.exe" -ArgumentList "$ArgumentList"
						Set-Location $Here
					}
					else
					{
						
						Write-log 'ADK is not present or you must Add components'
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						Write-log -message 'Bye..'
						Update-ListBox -Items $Log -Append -ListBox $listbox1
						
						
					}
				}
			}
			REFW10PRO.wim {
				
				Write-log -message "Extract the Windows 10 Professionnal index to a new WIM"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows 10 Professionnal" -DestinationImagePath $FPath
				
				Write-log -message "Create Mount Folder"
				[System.Windows.Forms.Application]::DoEvents()
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				if (!(Test-Path -path $MountFolder)) { New-Item -path $MountFolder -ItemType Directory | Out-Null }
				Write-log -message "Mount $RefImage Image..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Disable Protocol SMBv1"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				
				LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
				[System.Windows.Forms.Application]::DoEvents()
				Wait-Process -Name Dism
				[System.Windows.Forms.Application]::DoEvents()
				
				Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				add-WindowsPackage -PackagePath $Update -Path $MountFolder -PreventPending
				[System.Windows.Forms.Application]::DoEvents()
				Wait-Process -Name Dism
				
				Remove-AppxPro -MountFolder $MountFolder
				
				Write-log -message "Reduce size...."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Reduce size with Dism /Image:$MountFolder /Cleanup-Image /StartComponentCleanup /ResetBase"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "This Operation will take a long Time..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				
				[System.Windows.Forms.Application]::DoEvents()
				LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Cleanup-Image /StartComponentCleanup /ResetBase"
				Wait-Process -Name Dism
				Wait-Process -Name DismHost
				
				[System.Windows.Forms.Application]::DoEvents()
				
				Write-log -message "Add .NET Framework 3.5.1 to the Windows 10 Professionnal image "
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Dismount the Windows 10 Professionnal image"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				DisMount-WindowsImage -Path $MountFolder -Save
				[System.Windows.Forms.Application]::DoEvents()
				$Path2 = $Path +"\" + "REFW10ProInter.wim"
				Write-log -message "Export WimFile to reduce the size for Windows 10 Professionnal"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				Export-WindowsImage -SourceImagePath $FPath -CompressionType maximum -SourceName "Windows 10 Professionnal" -DestinationImagePath $Path2
				[System.Windows.Forms.Application]::DoEvents()
				
				Remove-Item -Path $FPath -Force
				[System.Windows.Forms.Application]::DoEvents()
				
				Rename-Item -Path "$Path2" -NewName "$RefImage"
				Write-log -message "Windows 10 Professionnal Image is finished"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				
				
			}
			REFW10EDU.wim {
				
				Write-log -message "Extract the Windows 10 Education index to a new WIM"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows 10 Education" -DestinationImagePath $FPath
				
				Write-log -message "Create Mount Folder"
				[System.Windows.Forms.Application]::DoEvents()
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				if (!(Test-Path -path $MountFolder)) { New-Item -path $MountFolder -ItemType Directory | Out-Null }
				Write-log -message "Mount $RefImage Image..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Disable Protocol SMBv1"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				
				LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
				[System.Windows.Forms.Application]::DoEvents()
				Wait-Process -Name Dism
				[System.Windows.Forms.Application]::DoEvents()
				
				Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				add-WindowsPackage -PackagePath $Update -Path $MountFolder -PreventPending
				[System.Windows.Forms.Application]::DoEvents()
				Wait-Process -Name Dism
				
				Remove-AppxPro -MountFolder $MountFolder
				
				Write-log -message "Reduce size...."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Reduce size with Dism /Image:$MountFolder /Cleanup-Image /StartComponentCleanup /ResetBase"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "This Operation will take a long Time..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				
				[System.Windows.Forms.Application]::DoEvents()
				LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Cleanup-Image /StartComponentCleanup /ResetBase"
				Wait-Process -Name Dism
				Wait-Process -Name DismHost
				
				[System.Windows.Forms.Application]::DoEvents()
				
				Write-log -message "Add .NET Framework 3.5.1 to the Windows 10 Education image "
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Dismount the Windows 10 Education image"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				DisMount-WindowsImage -Path $MountFolder -Save
				[System.Windows.Forms.Application]::DoEvents()
				$Path2 = $Path +"\"+ "REFW10EDUInter.wim"
				Write-log -message "Export WimFile to reduce the size for Windows 10 Education"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				Export-WindowsImage -SourceImagePath $FPath -CompressionType maximum -SourceName "Windows 10 Education" -DestinationImagePath $Path2
				[System.Windows.Forms.Application]::DoEvents()
				
				Remove-Item -Path $FPath -Force
				[System.Windows.Forms.Application]::DoEvents()
				
				Rename-Item -Path "$Path2" -NewName "$RefImage"
				Write-log -message "Windows 10 Education Image is finished"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				
				
				
				
			}
			REFW10ENT.wim {
				
				Write-log -message "Extract the Windows 10 Enterprise index to a new WIM"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Export-WindowsImage -SourceImagePath "$ISODrive\Sources\install.wim" -SourceName "Windows 10 Enterprise" -DestinationImagePath $FPath
				
				Write-log -message "Create Mount Folder"
				[System.Windows.Forms.Application]::DoEvents()
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				if (!(Test-Path -path $MountFolder)) { New-Item -path $MountFolder -ItemType Directory | Out-Null }
				Write-log -message "Mount $RefImage Image..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				Mount-WindowsImage -ImagePath $FPath -Index 1 -Path $MountFolder | Out-Null
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Disable Protocol SMBv1"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				
				LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Disable-Feature /FeatureName:SMB1Protocol"
				[System.Windows.Forms.Application]::DoEvents()
				Wait-Process -Name Dism
				[System.Windows.Forms.Application]::DoEvents()
				
				Write-log -message "Please be patient. Update the install.wim. This process takes around 20 to 30 minutes..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				add-WindowsPackage -PackagePath $Update -Path $MountFolder -PreventPending
				[System.Windows.Forms.Application]::DoEvents()
				Wait-Process -Name Dism
				
				Remove-AppxPro -MountFolder $MountFolder
				
				Write-log -message "Reduce size...."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Reduce size with Dism /Image:$MountFolder /Cleanup-Image /StartComponentCleanup /ResetBase"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "This Operation will take a long Time..."
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				
				[System.Windows.Forms.Application]::DoEvents()
				LaunchHide -Command "$env:SystemRoot\system32\Dism.exe" -arguments "/Image:$MountFolder /Cleanup-Image /StartComponentCleanup /ResetBase"
				Wait-Process -Name Dism
				Wait-Process -Name DismHost
				
				[System.Windows.Forms.Application]::DoEvents()
				
				Write-log -message "Add .NET Framework 3.5.1 to the Windows 10 Enterprise image "
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				Add-WindowsPackage -PackagePath $ISODrive\sources\sxs\microsoft-windows-netfx3-ondemand-package.cab -Path $MountFolder
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Dismount the Windows 10 Enterprise image"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				DisMount-WindowsImage -Path $MountFolder -Save
				[System.Windows.Forms.Application]::DoEvents()
				$Path2 = $Path +"\" + "REFW10ENTInter.wim"
				Write-log -message "Export WimFile to reduce the size for Windows 10 Enterprise"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				[System.Windows.Forms.Application]::DoEvents()
				
				Export-WindowsImage -SourceImagePath $FPath -CompressionType maximum -SourceName "Windows 10 Enterprise" -DestinationImagePath $Path2
				[System.Windows.Forms.Application]::DoEvents()
				
				Remove-Item -Path $FPath -Force
				[System.Windows.Forms.Application]::DoEvents()
				
				Rename-Item -Path "$Path2" -NewName "$RefImage"
				Write-log -message "Windows 10 Enterprise Image is finished"
				Update-ListBox -Items $Log -Append -ListBox $listbox1
				
				
			}
			default
			{
				#<code>
			}
		}
		# Dismount the Windows Server 2016 ISO
		[System.Windows.Forms.Application]::DoEvents()
		Dismount-DiskImage -ImagePath $SourceISO
		[System.Windows.Forms.Application]::DoEvents()
		$EndTime = Get-Date
		$duration = [math]::Round((New-TimeSpan -Start $StartTime -End $EndTime).TotalMinutes, 2)
		Write-log -message "========================================"
		Update-ListBox -Items $Log -Append -ListBox $listbox1
		Write-log -message "StartTime: $StartTime"
		Update-ListBox -Items $Log -Append -ListBox $listbox1
		Write-log -message "EndTime: $EndTime"
		Update-ListBox -Items $Log -Append -ListBox $listbox1
		Write-log -message "Duration: $duration minutes"
		Update-ListBox -Items $Log -Append -ListBox $listbox1
		Write-log -message "========================================"
		Update-ListBox -Items $Log -Append -ListBox $listbox1
	}
	
function LaunchHide ($Command, $arguments)
	{
		
		$startinfo = new-object System.Diagnostics.ProcessStartInfo
		$startinfo.FileName = $Command
		$startinfo.Arguments = $arguments
		$startinfo.CreateNoWindow = $true
		$startinfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::hidden
		$startinfo.UseShellExecute = $false
		[System.Diagnostics.Process]::Start($startinfo)
		
		
	}
	
function Remove-AppxPro ($MountFolder)
	{
		Write-log -message "=========Starting Step : Remove Win10 Apps========"
		Update-ListBox -Items $Log -Append -ListBox $listbox1
		[System.Windows.Forms.Application]::DoEvents()
		
		$apps = Get-AppxProvisionedPackage -Path $MountFolder
		$WhiteListedApps = @(
			"Microsoft.DesktopAppInstaller",
			"Microsoft.Messaging",
			"Microsoft.StorePurchaseApp"
			"Microsoft.WindowsCalculator",
			"Microsoft.WindowsCommunicationsApps", # Mail, Calendar etc
			"Microsoft.WindowsSoundRecorder",
			"Microsoft.WindowsStore"
		)
		[System.Windows.Forms.Application]::DoEvents()
		foreach ($i in $WhiteListedApps)
		{
			Write-log -message "The Application name $i is in the WhitelistedApps)"
			Update-ListBox -Items $Log -Append -ListBox $listbox1
			[System.Windows.Forms.Application]::DoEvents()
		}
		
		# Loop through the list of appx packages
		foreach ($App in $Apps)
		{
			# If application name not in appx package white list, remove AppxPackage and AppxProvisioningPackage
			if (($App.DisplayName -in $WhiteListedApps))
			{
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Skipping excluded application package: $($App.Displayname)"
				[System.Windows.Forms.Application]::DoEvents()
				
			}
			else
			{
				# Gather package names
				$AppPackageFullName = Get-AppxPackage -Name $App.DisplayName | Select-Object -ExpandProperty PackageFullName
				$AppProvisioningPackageName = Get-AppxProvisionedPackage -Path "$MountFolder" | Where-Object { $_.DisplayName -like $App.DisplayName } | Select-Object -ExpandProperty PackageName
				# Attempt to remove AppxPackage
				if ($AppPackageFullName -ne $null)
				{
					try
					{
						Write-log -message "Removing application package: $($App.Displayname)"
						[System.Windows.Forms.Application]::DoEvents()
						
						Remove-AppxProvisionedPackage -Path $MountFolder -PackageName $_.PackageName
						[System.Windows.Forms.Application]::DoEvents()
					}
					catch [System.Exception] {
						[System.Windows.Forms.Application]::DoEvents()
						#write-host  "Removing AppxPackage failed: $($_.Exception.Message)" -ForegroundColor DarkGreen
						Write-log -message "Removing AppxPackage failed: $($_.Exception.Message)"
						[System.Windows.Forms.Application]::DoEvents()
						
					}
				}
				else
				{
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Unable to locate AppxPackage for app: $($App.Displayname)"
					[System.Windows.Forms.Application]::DoEvents()
					
				}
				# Attempt to remove AppxProvisioningPackage
				if ($AppProvisioningPackageName -ne $null)
				{
					try
					{
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Removing application provisioning package: $($AppProvisioningPackageName)"
						[System.Windows.Forms.Application]::DoEvents()
						
						Remove-AppxProvisionedPackage -PackageName $AppProvisioningPackageName -Path $MountFolder -ErrorAction Stop | Out-Null
						[System.Windows.Forms.Application]::DoEvents()
					}
					catch [System.Exception] {
						[System.Windows.Forms.Application]::DoEvents()
						Write-log -message "Removing AppxProvisioningPackage failed: $($_.Exception.Message)"
						[System.Windows.Forms.Application]::DoEvents()
						
					}
				}
				else
				{
					[System.Windows.Forms.Application]::DoEvents()
					Write-log -message "Unable to locate AppxProvisioningPackage for app: $($App.Displayname)"
					
					[System.Windows.Forms.Application]::DoEvents()
				}
			}
		}
		# White list of Features On Demand V2 packages
		Write-log -message "Starting Features on Demand V2 removal process"
		Update-ListBox -Items $Log -Append -ListBox $listbox1
		
		$WhiteListOnDemand = "NetFX3|Tools.Graphics.DirectX|Tools.DeveloperMode.Core|Language|Browser.InternetExplorer"
		# Get Features On Demand that should be removed
		$OnDemandFeatures = Get-WindowsCapability -path $MountFolder | Where-Object { $_.Name -notmatch $WhiteListOnDemand -and $_.State -like "Installed" } | Select-Object -ExpandProperty Name
		foreach ($Feature in $OnDemandFeatures)
		{
			try
			{
				Write-log -message "Removing Feature on Demand V2 package: $($Feature)"
				[System.Windows.Forms.Application]::DoEvents()
				
				Get-WindowsCapability -Online -ErrorAction Stop | Where-Object { $_.Name -like $Feature } | Remove-WindowsCapability -Online -ErrorAction Stop | Out-Null
				[System.Windows.Forms.Application]::DoEvents()
			}
			catch [System.Exception] {
				[System.Windows.Forms.Application]::DoEvents()
				Write-log -message "Removing Feature on Demand V2 package failed: $($_.Exception.Message)"
				[System.Windows.Forms.Application]::DoEvents()
				
			}
		}
		
		Write-log -message "=========The step : Remove Win10 Apps is finished========"
		Update-ListBox -Items $Log -Append -ListBox $listbox1
	}