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
# Search for Roles
# Perform search of common Service Principal Names
#---------------------------------------------------

Function Get-Roles([String] $Filter)
	{
		$FinalObject = @()
		#Searcher
		if ($Global:Domain)
			{
				$Searcher = New-Object System.DirectoryServices.DirectorySearcher($Domain,$($credential.UserName),$($credential.GetNetworkCredential().password))
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
        if ($Filter)
        {
            $Searcher.SearchRoot ="LDAP://$($Filter)"
        }
		Foreach ($service in $Services)
		{
			$spn = [String]::Format($service.SPN)
			$Searcher.filter = "(servicePrincipalName=$spn*/*)"
        	$Results = $Searcher.FindAll()
			if ($Results)
				{
					Write-Verbose "Role Service Found....."
					foreach ($r in $Results)
						{
							$account = $r.GetDirectoryEntry()
							$record = @{}
							$record.Add("Name", [String]$account.Name)
							$record.Add("AccountName", [String]$account.sAMAccountName)
							$record.Add("GUID", [String]$account.guid)
							$record.Add("DN", [String]$account.DistinguishedName)
							$record.Add("Service", [String]$service.Service)
							$FinalObject +=$record
						}
				}
		}
	$AllRoles = Set-PSObject $FinalObject "Arsenal.AD.Roles"
	return $AllRoles
	}