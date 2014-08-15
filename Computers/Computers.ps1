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
# Search for Computers
#---------------------------------------------------

function Get-ComputersInfo ([String] $sAMAccountName, [Array] $Properties, [String] $Filter)
	{
        try
            {
	            $Computers = @()
	            # Filter on all Computer objects.
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
			            $Searcher.Filter = "(objectCategory=Computer)"
			            $Results = $Searcher.FindAll()
			            Write-Verbose "The computers search return $($Results.Count)" -Verbose
		            }
	            else
		            {
			            $Searcher.filter = "(&(objectClass=Computer)(sAMAccountName= $sAMAccountName))"
			            $Results = $Searcher.FindAll()
			            Write-Verbose "The computer search return $($Results.Count)" -Verbose
		            }
	            ForEach ($Result In $Results)
		            {
			            $record = @{}
			            ForEach ($Property in $Properties)
				            {
                                if ($Property -eq "lastLogon")
						            {
                                        if ([String]::IsNullOrEmpty($Result.Properties.Item([String]::Format($Property))))
                                            {
                                                $record.Add($Property,"Never")	
                                            }
                                        else
                                            {
							                    $lastLogon = [datetime]::FromFileTime([Int64]::Parse($Result.Properties.Item([String]::Format($Property))))
							                    $record.Add($Property,$lastLogon)
                                            }
						            }
                                elseif ($Property -eq "pwdLastSet")
						            {
                                        if ([String]::IsNullOrEmpty($Result.Properties.Item([String]::Format($Property))))
                                            {
                                                $record.Add($Property,0)
                                                	
                                            }
                                        else
                                            {
							                    $pwdLastSet = [datetime]::FromFileTime([Int64]::Parse($Result.Properties.Item([String]::Format($Property))))
							                    $record.Add($Property,$pwdLastSet)
                                                
                                            }
						            }
                                else
                                    {
					                    $record.Add($Property,[String] $Result.Properties.Item([String]::Format($Property)))	
                                    }
				            }
			    $Computers +=$record
			
		            }
            }
        catch
            {
                Write-Host "Error in Get-ComputersInfo function: $($_.Exception.Message)" -ForegroundColor Red
            }
	return $Computers
	}