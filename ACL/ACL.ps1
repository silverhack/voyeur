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
# Search ACL in domain
#---------------------------------------------------

function TranslateInheritanceType([String] $type)
	{
		$AppliesTo = ""
		$InheritanceType = @{"None"="This object only";"Descendents"="All child objects";`
							 "SelfAndChildren"="This object and one level Of child objects";`
							 "Children"="One level of child objects";"All"="This object and all child objects"}
		try
			{
				$AppliesTo = $InheritanceType[[String]::Format($type)]
			}
		catch
			{
				$AppliesTo = "Flag not resolved"
				Write-Debug $AppliesTo
			}
		return $AppliesTo
	}

function TranslateSid([Object] $SID)
	{
		$Account = ""
		$KnowSID = @{"S-1-0" ="Null Authority";"S-1-0-0"="Nobody";"S-1-1"="World Authority";"S-1-1-0"="Everyone";`
					 "S-1-2"="Local Authority";"S-1-2-0"="Local";"S-1-2-1"="Console Logon";"S-1-3"="Creator Authority";`
					 "S-1-3-0"="Creator Owner";"S-1-3-1"="Creator Group";"S-1-3-4"="Owner Rights";"S-1-5-80-0"="All Services";`
					 "S-1-4"="Non Unique Authority";"S-1-5"="NT Authority";"S-1-5-1"="Dialup";"S-1-5-2"="Network";`
					 "S-1-5-3"="Batch";"S-1-5-4"="Interactive";"S-1-5-6"="Service";"S-1-5-7" ="Anonymous";"S-1-5-9"="Enterprise Domain Controllers";`
					 "S-1-5-10"="Self";"S-1-5-11"="Authenticated Users";"S-1-5-12"="Restricted Code";"S-1-5-13"="Terminal Server Users";`
					 "S-1-5-14"="Remote Interactive Logon";"S-1-5-15"="This Organization";"S-1-5-17"="This Organization";"S-1-5-18"="Local System";`
					 "S-1-5-19"="NT Authority Local Service";"S-1-5-20"="NT Authority Network Service";"S-1-5-32-544"="Administrators";`
					 "S-1-5-32-545"="Users";"S-1-5-32-546"="Guests";"S-1-5-32-547"="Power Users";"S-1-5-32-548"="Account Operators";`
					 "S-1-5-32-549"="Server Operators";"S-1-5-32-550"="Print Operators";"S-1-5-32-551"="Backup Operators";"S-1-5-32-552"="Replicators";`
					 "S-1-5-32-554"="Pre-Windows 2000 Compatibility Access";"S-1-5-32-555"="Remote Desktop Users";"S-1-5-32-556"="Network Configuration Operators";`
					 "S-1-5-32-557"="Incoming forest trust builders";"S-1-5-32-558"="Performance Monitor Users";"S-1-5-32-559"="Performance Log Users";`
					 "S-1-5-32-560"="Windows Authorization Access Group";"S-1-5-32-561"="Terminal Server License Servers";"S-1-5-32-562"="Distributed COM Users";`
					 "S-1-5-32-569"="Cryptographic Operators";"S-1-5-32-573"="Event Log Readers";"S-1-5-32-574"="Certificate Services DCOM Access";`
					 "S-1-5-32-575"="RDS Remote Access Servers";"S-1-5-32-576"="RDS Endpoint Servers";"S-1-5-32-577"="RDS Management Servers";`
					 "S-1-5-32-578"="Hyper-V Administrators";"S-1-5-32-579"="Access Control Assistance Operators";"S-1-5-32-580"="Remote Management Users"}
		try
			{			
				$Account = $SID.Translate([System.Security.Principal.NTAccount])
				#Write-Host $Account -ForegroundColor Red
			}
		catch
			{
				$Account = $KnowSID[[String]::Format($SID)]	
				#Write-Host $Account -ForegroundColor Yellow
			}
		return $Account
		
	}
	
function GetACL([String] $object, [Bool]$full, [String] $Filter)
	{
		if ($Global:Domain)
		{
            #$D = [adsi]"LDAP://$($Filter)"
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
            $Searcher.SearchRoot = "LDAP://$($Filter)"
        }
           $Searcher.filter = $object
		        $Searcher.PropertiesToLoad.AddRange("DistinguishedName") > $null
		        $Searcher.PropertiesToLoad.AddRange("name") > $null
		        $Results = $Searcher.FindAll()
		if($full)
			{
				$ListOU = @()
				foreach ($result in $Results)
					{
						#$ListOU += $result.properties.item("DistinguishedName")
						$data = $result.GetDirectoryEntry()
						$aclObject = $data.Get_ObjectSecurity()
						$aclList = $aclObject.GetAccessRules($true,$true,[System.Security.Principal.SecurityIdentifier])
						foreach($acl in $aclList)
         					{
             					$objSID = New-Object System.Security.Principal.SecurityIdentifier($acl.IdentityReference)
			 					$AccountName = TranslateSID($objSID)
			 					$AppliesTo = TranslateInheritanceType($acl.InheritanceType)
             					$OUACL = @{
			 								"AccountName" = $AccountName
											"DN" = [String]::Format($result.properties.item('DistinguishedName'))
											"Rights" = $acl.ActiveDirectoryRights
											"Risk" = Get-Icon $acl.ActiveDirectoryRights
											"AppliesTo" = $AppliesTo
											"AccessControlType" = $acl.AccessControlType
											"InheritanceFlags" = $acl.InheritanceFlags
								  		}           
            					$obj = New-Object -TypeName PSObject -Property $OUACL
            					$obj.PSObject.typenames.insert(0,'Arsenal.AD.OUACL')
								$ListOU +=$obj
            					#Write-Output $obj		
							}
					}
				return $ListOU
			}
		else
			{
				$ListOU = @()
				foreach ($result in $Results)
					{
						$OUProperties = @{
									"Name" = [String]::Format($result.properties.item('name'))
									"DN" = [String]::Format($result.properties.item('DistinguishedName'))
								}
						$obj = New-Object -TypeName PSObject -Property $OUProperties
            			$obj.PSObject.typenames.insert(0,'Arsenal.AD.OUProperties')
						$ListOU +=$obj
					}
				return $ListOU				
			}		
	}

