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
# Search for High privileges account
#---------------------------------------------------

function Get-HighPrivileges
    {
        #http://support.microsoft.com/kb/243330
        $PrivGroups = @("S-1-5-32-552","S-1-5-32-544";"S-1-5-32-548";"S-1-5-32-549";"S-1-5-32-551";`
                        "$RootDomainSID-519";"$RootDomainSID-518";"$DomainSID-512";"$DomainSID-521")
        $HighPrivileges = @()
		if ($Global:Domain)
		{
			$Searcher = New-Object System.DirectoryServices.DirectorySearcher($Domain)
			$Searcher.SearchScope = "subtree"
			$Searcher.PageSize = 200
		}
		elseif ($Global:DomainWPass)
			{
				$Searcher = New-Object System.DirectoryServices.DirectorySearcher($DomainWPass)
				$Searcher.SearchRoot = "LDAP://" + $DomainWPass
				$Searcher.PageSize = 200
			}
		else
		{
			$Searcher = New-Object System.DirectoryServices.DirectorySearcher($CurrentDomain.distinguishedname)
			$Searcher.SearchRoot = "LDAP://" + $CurrentDomain.distinguishedname		
			$Searcher.PageSize = 200
		}
        $ListPrivGroups = @{}
		foreach ($group in $PrivGroups)
			{
			$Searcher.filter = "(objectSID=$group)"
        	$Results = $Searcher.FindOne()
			ForEach ($result in $Results)
        		{
            	$dn = $result.properties.distinguishedname
				$cn = $result.properties.cn   
            	$ListPrivGroups.Add($cn,$dn)
            	}
			}
        #$HighPrivileges = @{}
		$ListPrivGroups.GetEnumerator() | ForEach-Object `
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
            try
                {
                    $ErrorActionPreference = "Continue"
			        $Searcher.PageSize = 200
			        $Searcher.Filter = "(&(objectCategory=group)(name=$($_.Key)))"
			        $Items = $Searcher.FindAll()
                }
            catch
                {
                    $ErrorActionPreference = "Continue"
                    Write-Host "Function Get-HighPrivileges. Error in group $($_.Key)" -ForegroundColor Red
                }
            
			foreach ($Item in $Items)
				{
					$Group = $Item.GetDirectoryEntry()
					$Members = $Group.member				
				}
			try
				{
				if (($Group.Member).Count -gt 0)
					{
					foreach	($Member in $Members)
						{
						$Properties = @("pwdLastSet","whenCreated")
						$sAMAccountName = ($Member.replace('\','') -split ",*..=")[1] 
						if ($Domain)
							{
								$User = Get-UsersInfo -Properties $Properties -sAMAccountName $sAMAccountName
							}
						$HighPrivilege = @{
										"AccountName" = [String]::Format($sAMAccountName)
										"DN" = [String]::Format($Member)
										"MemberOf" = [String]::Format($_.Value)
										"PasswordLastSet" = $User.pwdLastSet
										"whenCreated" = $User.whenCreated
										"Group" = [String]::Format($_.Key)
										}
						$obj = New-Object -TypeName PSObject -Property $HighPrivilege
            			$obj.PSObject.typenames.insert(0,'Arsenal.AD.HighPrivileges')
            			$HighPrivileges += $obj
						}
					}		
				}
			catch
				{
					Write-Host "Function Get-HighPrivileges $($_.Exception.Message)" -ForegroundColor Red
				}
			}
		return $HighPrivileges
    }

<#
#---------------------------------------------------
# Search for High privileges account
#---------------------------------------------------
function Get-HighPrivileges( $Grupitos)
	{
		$HighPrivileges = @()
		if ($Global:Domain)
		{
			$Searcher = New-Object System.DirectoryServices.DirectorySearcher($Domain)
			$Searcher.SearchScope = "subtree"
			$Searcher.PageSize = 200
		}
		elseif ($Global:DomainName)
			{
				$Searcher = New-Object System.DirectoryServices.DirectorySearcher($DomainName)
				$Searcher.SearchRoot = "LDAP://" + $DomainName
				$Searcher.PageSize = 200
			}
		else
		{
			$Searcher = New-Object System.DirectoryServices.DirectorySearcher($CurrentDomain.distinguishedname)
			$Searcher.SearchRoot = "LDAP://" + $CurrentDomain.distinguishedname		
			$Searcher.PageSize = 200
		}
		#$Connect = Get-CurrentDomain
		#$Connect = [ADSI]"LDAP://$($Domain.distinguishedName)"
		$ListPrivGroups = @{}
		foreach ($grupo in $Grupitos)
			{
			$Searcher.filter = "(CN=$grupo)"
        	$Results = $Searcher.FindAll()
			ForEach ($result in $Results)
        		{
            	$dn = $result.properties.distinguishedname
				$cn = $result.properties.cn
            	$ListPrivGroups.Add($cn,$dn)
            	}
			}
		#$HighPrivileges = @{}
		$ListPrivGroups.GetEnumerator() | ForEach-Object `
			{
			if ($Global:Domain)
				{
					$Searcher = New-Object System.DirectoryServices.DirectorySearcher($Domain)
					$Searcher.SearchScope = "subtree"
				}
			elseif ($Global:DomainName)
				{
					$Searcher = New-Object System.DirectoryServices.DirectorySearcher($DomainName)
					$Searcher.SearchRoot = "LDAP://" + $DomainName
				}
			else
				{
					$Searcher = New-Object System.DirectoryServices.DirectorySearcher($CurrentDomain.distinguishedname)
					$Searcher.SearchRoot = "LDAP://" + $CurrentDomain.distinguishedname		
				}
			$Searcher.PageSize = 200
			$Searcher.Filter = "(&(objectCategory=group)(name=$($_.Key)))"
			$Items = $Searcher.FindAll()
			foreach ($Item in $Items)
				{
					$Group = $Item.GetDirectoryEntry()
					$Members = $Group.member				
				}
			try
				{
				if (($Group.Member).Count -gt 0)
					{
					foreach	($Member in $Members)
						{
						$Properties = @("pwdLastSet","whenCreated")
						$sAMAccountName = ($Member.replace('\','') -split ",*..=")[1] 
						if ($Domain)
							{
								$User = Get-UsersInfo -Properties $Properties -sAMAccountName $sAMAccountName
							}
						$HighPrivilege = @{
										"AccountName" = [String]::Format($sAMAccountName)
										"DN" = [String]::Format($Member)
										"MemberOf" = [String]::Format($_.Value)
										"PasswordLastSet" = $User.pwdLastSet
										"whenCreated" = $User.whenCreated
										"Group" = [String]::Format($_.Key)
										}
						$obj = New-Object -TypeName PSObject -Property $HighPrivilege
            			$obj.PSObject.typenames.insert(0,'Arsenal.AD.HighPrivileges')
            			$HighPrivileges += $obj
						}
					}		
				}
			catch
				{
					Write-Host "$($_.Exception.Message)" -ForegroundColor Red
				}
			}
		return $HighPrivileges
	}
#>
<#
$tmp = Get-CurrentDomain
$data = Get-HighPrivileges($tmp)
$data | Format-Table
#>

