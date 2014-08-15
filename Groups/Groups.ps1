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
# Search for groups
#---------------------------------------------------
function Get-GroupInfo ([String] $sAMAccountName, [Array] $Properties, [String] $Filter)
	{
		$Groups = @()
		# Filter on all groups objects.
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
		# Attribute values to retrieve.
		foreach ($property in $Properties)
			{
			$Searcher.PropertiesToLoad.Add([String]::Format($property)) > $Null
			}
		if (!$sAMAccountName)
		{
			$Searcher.Filter = "(&(objectCategory=group)(objectClass=group))"
			$Results = $Searcher.FindAll()
			#Write-Verbose "The group search return $($Results.Count)" -Verbose
		}
		else
		{
			$Searcher.filter = "(&(objectClass=group)(sAMAccountName= $($sAMAccountName)))"
			$Results = $Searcher.FindAll()
			#Write-Verbose "The user search return $($Results.Count)" -Verbose
		}
		foreach ($result in $Results)
			{
				$record = @{}
				ForEach ($Property in $Properties)
				{
					if ($property -eq "groupType")
						{
							[String]$tmp = $Result.Properties.Item($Property)
							$record.Add($Property,$GroupType[$tmp])							
						}
					else
						{
							$record.Add($Property,[String]$Result.Properties.Item([String]::Format($Property)))
						}
					
				}
			$Groups +=$record
			}
		return $Groups		
	}