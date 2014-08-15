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
# Construct detailed PSObject
#---------------------------------------------------

function Set-PSObject ([Object] $Object, [String] $type)
	{
		$FinalObject = @()
		foreach ($Obj in $Object)
			{
				$NewObject = New-Object -TypeName PSObject -Property $Obj
            	$NewObject.PSObject.typenames.insert(0,[String]::Format($type))
				$FinalObject +=$NewObject
			}
		return $FinalObject
	}

#---------------------------------------------------
# Get Color of ACL
#---------------------------------------------------

Function Get-Color ([String] $Color)
	{
		try
			{
				return $ACLColors[$Color]
			}
		catch
			{
				Write-Host "ACL color not found" -ForegroundColor Red
			}
	}

#---------------------------------------------------
# Get Icon of ACL
#---------------------------------------------------

Function Get-Icon ([String] $Color)
	{
		try
			{
				return $ACLIcons[$Color]
			}
		catch
			{
				Write-Host "ACL color not found" -ForegroundColor Red
			}
	}

#---------------------------------------------------
# Volumetry Organizational Unit
#---------------------------------------------------
function Get-OUVolumetry ([Object] $Results, [Object] $FinalUsers)
	{
 
		if ($Global:Domain)
		{
			$Searcher = New-Object System.DirectoryServices.DirectorySearcher($Domain)
			$Searcher.SearchScope = "subtree"
		}
		elseif ($Global:DomainWPass)
		{
			$Searcher = New-Object System.DirectoryServices.DirectorySearcher($DomainWPass)
			$Searcher.SearchRoot = "LDAP://" + $DomainWPass
		}
		else
		{
			$Searcher = New-Object System.DirectoryServices.DirectorySearcher($CurrentDomain.distinguishedname)
			$Searcher.SearchRoot = "LDAP://" + $CurrentDomain.distinguishedname		
		}
		$Searcher.PageSize = 200
		$ListOU = @()
		foreach ($result in $Results)
			{
				$ou =[String]::Format($result.DN)
				$ouname =[String]::Format($result.Name)
				$Searcher.filter = "(&(objectclass=organizationalunit)(distinguishedname=$ou))"
				$tmpou = $Searcher.FindOne()
				$data = $tmpou.GetDirectoryEntry()
				$NumberOfUsers = 0
				if ($data)
					{
						$NumberOfUsers = ($data.Children | Where-Object {$_.schemaClassName -eq "user"}).Count
						$TotalUsers = $data.Children | Where-Object {$_.schemaClassName -eq "user"} | Select-Object sAMAccountName
						$InactiveUsers = 0
						$ActiveUsers = 0
						foreach ($u in $TotalUsers)
							{
								Foreach	($user in $FinalUsers)
									{
										if ($user.sAMAccountName -eq $u.sAMAccountName -and $user.isActive -eq $true)
											{
												$ActiveUsers++
											}
										if ($user.sAMAccountName -eq $u.sAMAccountName -and $user.isActive -eq $false)
											{
												$InactiveUsers++
											}
									}
							}
					}
                    try
                        {
                            
					            $OuCount = @{
					 		    "Name" = $ouname
					 		    "distinguishedName" = $ou
					 		    "NumberOfUsers" = $NumberOfUsers
							    "ActiveUsers" = $ActiveUsers
							    "InactiveUsers" = $InactiveUsers
							    }
					            $obj = New-Object -TypeName PSObject -Property $OuCount
        			            $obj.PSObject.typenames.insert(0,'Arsenal.AD.OuCount')
					            $ListOU +=$obj	
                                #Need to resolv filters
                        }
                    catch
                        {
                            Write-Host "Get-OUVolumetry problem...$($_.Exception.Message) in $($NumberofUsers)" -ForegroundColor Red
                        }
			}
		return $ListOU
	}
	
function New-Report([String] $Path, [String] $Domain)
	{
		$target = "$Path\Reports"
		if (!(Test-Path -Path $target))
			{
				$tmpdir = New-Item -ItemType Directory -Path $target
				Write-Host "Folder Reports created in $target...." -ForegroundColor Yellow
			}
		$folder = "$target\" + ([System.Guid]::NewGuid()).ToString() + $Domain
		if (!(Test-Path -Path $folder))
			{
				try
					{
						$tmpdir = New-Item -ItemType Directory -Path $folder
						return $folder
					}
				catch
					{
						Write-Verbose "Failed to create new directory. Trying to generate new guid...."
						New-Report -Path $Path -Domain $Domain
					}
			}
				
	}