<#
.SYNOPSIS
	The script scan Active Directory and retrieve huge data of many sources like:
	-User objects
	-Computer objects
	-Organizational Unit
	-Forest
	-Access control list
	-Etc.
	
	Ability to export data to CSV and Excel format. Default save format is CSV
	
.NOTES
	Author		: Juan Garrido (http://windowstips.wordpress.com)
	Twitter		: @tr1ana
	Company		: http://www.innotecsystem.com
	File Name	: voyeur.ps1	

#>

#---------------------------------------------------
# General Status of Active Directory objects
#--------------------------------------------------- 

function Get-UserStatus([Object]$Users)
	{
		for ($i=0; $i -lt $Users.Count; $i++)
			{
				Write-Progress -activity "Collecting Users metrics..." -status "$($DomainNameString): processing of $i/$($Users.Count) Users" -PercentComplete (($i / $Users.Count) * 100)
				# Active accounts
				if ( ([string]::IsNullOrEmpty($Users[$i].lastlogontimestamp)) -or (([DateTime]$Users[$i].lastlogontimestamp).ticks -lt $InactiveDate))
						{
							$Users[$i].isActive = $false
							$UsersCount['isInactive']['Count']++	
						}
						else
						{
							$Users[$i].isActive = $true
							$UsersCount['isActive']['Count']++
							
						}
					#Password Expires
				if ($Users[$i].ADS_UF_PASSWORD_EXPIRED -eq $true)
					{
						if ($Users[$i].isActive)
							{
								$UsersCount['isActive']['isPwdExpired']++
							}
						else
							{
								$UsersCount['isInactive']['isPwdExpired']++
							}
					}
				if ($Users[$i].accountExpires -lt $CurrentDate)
					{
						$Users[$i].isExpired = $true
						if ($Users[$i].isActive)
							{
								$UsersCount['isActive']['isExpired']++
							}
						else
							{
								$UsersCount['isInactive']['isExpired']++
							}
					}
				else
					{
						$Users[$i].isExpired = $false						
					}
				if ($Users[$i].isDisabled)
					{
						if ($Users[$i].isActive)
							{
								$UsersCount['isActive']['isDisabled']++
							}
						else
							{
								$UsersCount['isInactive']['isDisabled']++
							}
					}
				if ($Users[$i].isPwdNoRequired)
					{
						if ($Users[$i].isActive)
							{
								$UsersCount['isActive']['isPwdNotRequired']++
							}
						else
							{
								$UsersCount['isInactive']['isPwdNotRequired']++
							}
					}
				if ($Users[$i].isNoDelegation -eq $false)
					{
						if ($Users[$i].isActive)
							{
								$UsersCount['isActive']['isPTH']++
							}
						else
							{
								$UsersCount['isInactive']['isPTH']++
							}
					}
				# Identify pwdLastSet older than $PasswordAge days
						if ( $Users[$i].pwdLastSet -lt $PasswordLastSet )
						{
							$Users[$i].isPwdOld = $true
							
							if ( $UserObj.isActive )
							{
								$UsersCount['isActive']['isPwdOld']++
							}
							else
							{
								$UsersCount['isInactive']['isPwdOld']++
							}
						}
				# Identify password never expires
						if ( $Users[$i].isPwdNeverExpires)
						{							
							if ( $Users[$i].isActive )
							{
								$UsersCount['isActive']['isPwdNeverExpires']++
							}
							else
							{
								$UsersCount['isInactive']['isPwdNeverExpires']++
							}
						}
			}
		return $Users
	}
	

#---------------------------------------------------
# General Status of Active Directory objects
#--------------------------------------------------- 

function Get-ComputerStatus([Object]$Computers)
	{
		for ($i=0; $i -lt $Computers.Count; $i++)
			{
				Write-Progress -activity "Collecting Computers metrics..." -status "$($DomainNameString): processing of $i/$($Computers.Count) Computers" -PercentComplete (($i / $Computers.Count) * 100)
				# Active accounts
				if ($Computers[$i].operatingSystem -eq "Windows XP Professional" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 3")
					{
						$ComputersCount['WindowsXPSP3']++
					}
				if ($Computers[$i].operatingSystem -eq "Windows XP Professional" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 2")
					{
						$ComputersCount['WindowsXPSP2']++
					}
				if ($Computers[$i].operatingSystem -eq "Windows XP Professional" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 1")
					{
						$ComputersCount['WindowsXPSP1']++
					}
				if ($Computers[$i].operatingSystem -eq "Windows Server 2008 R2 Enterprise" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 1")
					{
						$ComputersCount['WindowsServer2008R2SP1']++
					}
				if ($Computers[$i].operatingSystem -eq "Windows 8 Professional" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 1")
					{
						$ComputersCount['Windows8']++
					}
                if ($Computers[$i].operatingSystem -eq "Windows 8 Pro" -and $Computers[$i].operatingSystemServicePack -eq $null)
					{
						$ComputersCount['Windows8Pro']++
					}
                if ($Computers[$i].operatingSystem -eq "Windows 8.1 Pro" -and $Computers[$i].operatingSystemServicePack -eq $null)
					{
						$ComputersCount['Windows81Pro']++
					}
				if ($Computers[$i].operatingSystem -eq "Windows 7 Professional" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 1")
					{
						$ComputersCount['Windows7ProfessionalSP1']++
					} 
                if ($Computers[$i].operatingSystem -eq "Windows 7 Professional" -and $Computers[$i].operatingSystemServicePack -eq $null)
					{
						$ComputersCount['Windows7Professional']++
					} 
                if ($Computers[$i].operatingSystem -eq "Windows 7 Ultimate" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 1")
					{
						$ComputersCount['Windows7UltimateSP1']++
					}  
                 if ($Computers[$i].operatingSystem -eq "Windows 7 Enterprise" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 1")
					{
						$ComputersCount['Windows7EnterpriseSP1']++
					}           
                if ($Computers[$i].operatingSystem -eq "Windows Server 2003" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 2")
					{
						$ComputersCount['Windows2003SP2']++
                        
					}
                if ($Computers[$i].operatingSystem -eq "Windows Server 2003" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 1")
					{
						$ComputersCount['Windows2003SP1']++
					}
                if ($Computers[$i].operatingSystem -eq "Windows Server 2003" -and $Computers[$i].operatingSystemServicePack -eq $null)
					{
						$ComputersCount['Windows2003']++
					}
                if ($Computers[$i].operatingSystem -eq "Windows 2000 Server" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 4")
					{
						$ComputersCount['Windows2000SP4']++
					}
                if ($Computers[$i].operatingSystem -eq "Windows Server 2008 Standard" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 1")
					{
						$ComputersCount['WindowsServer2008STDSP1']++
					}
                if ($Computers[$i].operatingSystem -eq "Windows Server 2008 Standard" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 2")
					{
						$ComputersCount['WindowsServer2008STDSP2']++
					}
                if ($Computers[$i].operatingSystem -eq "Windows Server 2008 R2 Standard" -and $Computers[$i].operatingSystemServicePack -eq $null)
					{
						$ComputersCount['WindowsServer2008R2STD']++
					}
                if ($Computers[$i].operatingSystem -eq "Windows Server 2008 R2 Standard" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 1")
					{
						$ComputersCount['WindowsServer2008R2STDSP1']++
					}
                if ($Computers[$i].operatingSystem -eq "Windows Server 2008 Enterprise" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 1")
					{
						$ComputersCount['WindowsServer2008SP1']++
					}
                if ($Computers[$i].operatingSystem -eq "Windows Server 2008 Enterprise" -and $Computers[$i].operatingSystemServicePack -eq "Service Pack 2")
					{
						$ComputersCount['WindowsServer2008SP2']++
					}
                if ($Computers[$i].operatingSystem -eq "Windows Storage Server 2012 Standard" -and $Computers[$i].operatingSystemServicePack -eq $null)
					{
						$ComputersCount['WindowsServer2012STDStorage']++
					}
                if ($Computers[$i].operatingSystem -eq "SLES" -and $Computers[$i].operatingSystemServicePack -contains "Likewise")
					{
						$ComputersCount['LikeWiseOpen']++
					}
                if ($Computers[$i].operatingSystem -eq "Samba" -and $Computers[$i].operatingSystemServicePack -eq $null)
					{
						$ComputersCount['Samba']++
					}
                if ($Computers[$i].operatingSystem -eq "EMC File Server" -and $Computers[$i].operatingSystemServicePack -eq $null)
					{
						$ComputersCount['EMCFileServer']++
					}
			}
	}
				