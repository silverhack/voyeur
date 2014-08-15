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
# Return FQDN
#---------------------------------------------------
Function FQDN2DN
{
	Param ($domainFQDN)
	$colSplit = $domainFQDN.Split(".")
	$FQDNdepth = $colSplit.length
	$DomainDN = ""
	For ($i=0;$i -lt ($FQDNdepth);$i++)
	{
		If ($i -eq ($FQDNdepth - 1)) {$Separator=""}
		else {$Separator=","}
		[string]$DomainDN += "DC=" + $colSplit[$i] + $Separator
	}
	$DomainDN
}

#---------------------------------------------------
# Return the Root Domain SID
#---------------------------------------------------

function Get-RootDomainSID
    {
        try
            {
                $ForestRootDN = FQDN2DN $Forest.rootdomain.name
                $GlobalCatalog = $Forest.FindGlobalCatalog()
                $DN = "GC://$($GlobalCatalog.IPAddress)/$($ForestRootDN)"
				if ($Global:Domain)
		    		{
                	$ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($DN,$($credential.UserName),$($credential.GetNetworkCredential().password))
					}
				else
					{
						$ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($DN)
					}
                #$ADObject | gm
                $RootDomainSid = New-Object System.Security.Principal.SecurityIdentifier($ADObject.objectSid[0], 0)
                return $RootDomainSid.ToString()
            }
        catch
            {
                Write-Host "Function Get-RootDomainSID $($_.Exception.Message)" -ForegroundColor Red
            }     
        
    }


#---------------------------------------------------
# Return the Domain SID
#---------------------------------------------------

function Get-DomainSID
    {
        try
            {
                $GlobalCatalog = $Forest.FindGlobalCatalog()
				if ($Global:Domain)
		    		{
					$DN = "GC://$($GlobalCatalog.IPAddress)/$($Domain.distinguishedname)"
                	$ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($DN,$($credential.UserName),$($credential.GetNetworkCredential().password))
					}
				elseif ($Global:DomainWPass)
					{
						$DN = "GC://$($GlobalCatalog.IPAddress)/$($DomainWPass.distinguishedname)"
						$ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($DN)
					}
				else
					{
						$DN = "GC://$($GlobalCatalog.IPAddress)/$($CurrentDomain.distinguishedname)"
						$ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($DN)
					}
                #$ADObject | gm
                $DomainSid = New-Object System.Security.Principal.SecurityIdentifier($ADObject.objectSid[0], 0)
                return $DomainSid.ToString()
            }
        catch
            {
                Write-Host "Function Get-DomainSID $($_.Exception.Message)" -ForegroundColor Red
            }     
        
    }

#---------------------------------------------------
# Return the Current Forest
#---------------------------------------------------

function Get-CurrentForest
    {
		try
			{
		        if ($Global:Domain)
				    {
		            # Create domain level connection
		            $context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$DomainPassed,$($credential.UserName),$($credential.GetNetworkCredential().password))
		            $MyDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($context)
            
		            }
		        elseif ($Global:DomainWPass)
					{
		             # Create domain level connection
		            $context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$DomainWPass.name)
		            $MyDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($context)
		            }
		        else
				{
		             $forest = [System.DirectoryServices.ActiveDirectory.forest]::getcurrentforest()
		        }

				if ($Global:Domain)
				    {
		        		#Forest Level Connection
		        		$ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest", $MyDomain.forest,$($credential.UserName),$($credential.GetNetworkCredential().password))
		        		$forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
					}
                elseif ($Global:DomainWPass)
                    {
                        #Forest Level Connection
		        		$ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest", $MyDomain.forest)
		        		$forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
                    }
			}
		Catch
			{
				Write-Host "Function Get-CurrentForest $($_.Exception.Message)" -ForegroundColor Red
			}
        return $forest

    }

#---------------------------------------------------
# Return the current Domain
#---------------------------------------------------
function Get-CurrentDomain
	{
		try
			{
				$current = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
				$Domain = [ADSI]"LDAP://$current"
				Write-Host "Establishing connection with domain: " -NoNewline
				Write-Host $Domain.name -ForegroundColor Magenta
				return $Domain 
			}
		catch
			{
				Write-host "unable to contact..." -ForegroundColor Red
			}
	}

#---------------------------------------------------
# Try to return domain passed variable
#---------------------------------------------------
function Get-Domain ([String] $current)
	{
		try
			{
				$Domain = [ADSI]"LDAP://$current"
				Write-Host "Establishing connection with domain: " -NoNewline
				Write-Host $Domain.name -ForegroundColor Magenta
				$Searcher = New-Object System.DirectoryServices.DirectorySearcher
				$Searcher.SearchScope = "subtree"
				return $Domain
			}
		catch
			{
				Write-host "unable to contact..." -ForegroundColor Red
			}
	}

#---------------------------------------------------
# Try to return domain with alternate credentials
#---------------------------------------------------
function Get-AuthDomain ([String] $current, [String] $username)
	{
		try
			{
				$DN = "LDAP://"+$current
				#Create domain object
				if ($UseSSL -eq $true)
					{
						$Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($DN,$($credential.UserName),$($credential.GetNetworkCredential().password), "SecureSocketsLayer")
					}
				else
					{
						$Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($DN,$($credential.UserName),$($credential.GetNetworkCredential().password))
					}
				Write-Host "Establishing connection with domain: " -NoNewline
				Write-Host $Domain.name -ForegroundColor Magenta
				<#
				$Searcher.PageSize = 200
				$Searcher.SearchRoot= $Domain
				$Searcher.Filter = "(&(objectCategory=person)(objectClass=user))"
				$Results = $Searcher.FindAll()
				Write-Verbose "The users search return $($Results.Count)" -Verbose
				#>
				return $Domain
			}
		catch
			{
				Write-host "unable to contact..." -ForegroundColor Red
			}
	}