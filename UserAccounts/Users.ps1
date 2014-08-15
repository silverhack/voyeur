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

function Set-PSObjectUser ([Object] $Users)
	{
		$FinalUsers = @()
		foreach ($User in $Users)
			{
				$obj = New-Object -TypeName PSObject -Property $User
            	$obj.PSObject.typenames.insert(0,'Arsenal.AD.Users')
				$FinalUsers +=$obj
			}
		return $FinalUsers
	}

#---------------------------------------------------
# Resolv Bits of msds-user-account-control-computed
#---------------------------------------------------

function Get-MsDSUACResolv ([Object] $Users)
	{
        Write-Host "Resolving Attributes..." -ForegroundColor Green
		$FinalUsers = @()
		$ADS_USER_FLAG_ENUM = @{ 
  					ADS_UF_SCRIPT                                  = 1;        #// 0x1
  					ADS_UF_ACCOUNTDISABLE                          = 2;        #// 0x2
  					ADS_UF_HOMEDIR_REQUIRED                        = 8;        #// 0x8
  					ADS_UF_LOCKOUT                                 = 16;       #// 0x10
  					ADS_UF_PASSWD_NOTREQD                          = 32;       #// 0x20
  					ADS_UF_PASSWD_CANT_CHANGE                      = 64;       #// 0x40
  					ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED         = 128;      #// 0x80
  					ADS_UF_TEMP_DUPLICATE_ACCOUNT                  = 256;      #// 0x100
  					ADS_UF_NORMAL_ACCOUNT                          = 512;      #// 0x200
  					ADS_UF_INTERDOMAIN_TRUST_ACCOUNT               = 2048;     #// 0x800
  					ADS_UF_WORKSTATION_TRUST_ACCOUNT               = 4096;     #// 0x1000
  					ADS_UF_SERVER_TRUST_ACCOUNT                    = 8192;     #// 0x2000
  					ADS_UF_DONT_EXPIRE_PASSWD                      = 65536;    #// 0x10000
  					ADS_UF_MNS_LOGON_ACCOUNT                       = 131072;   #// 0x20000
  					ADS_UF_SMARTCARD_REQUIRED                      = 262144;   #// 0x40000
  					ADS_UF_TRUSTED_FOR_DELEGATION                  = 524288;   #// 0x80000
  					ADS_UF_NOT_DELEGATED                           = 1048576;  #// 0x100000
  					ADS_UF_USE_DES_KEY_ONLY                        = 2097152;  #// 0x200000
  					ADS_UF_DONT_REQUIRE_PREAUTH                    = 4194304;  #// 0x400000
  					ADS_UF_PASSWORD_EXPIRED                        = 8388608;  #// 0x800000
  					ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION  = 16777216; #// 0x1000000
  		}
	$Attribute = "msds-user-account-control-computed"
	if ($Global:Domain)
		{
			$Searcher = New-Object System.DirectoryServices.DirectorySearcher($Domain)
		}
	elseif ($Global:DomainWPass)
		{
			$Searcher = New-Object System.DirectoryServices.DirectorySearcher($DomainWPass)
		}
	else
		{
			$Searcher = New-Object System.DirectoryServices.DirectorySearcher($CurrentDomain.distinguishedname)
			$Searcher.SearchRoot = "LDAP://" + $CurrentDomain.distinguishedname		
		}
	try
		{
		foreach ($User in $Users)
			{
			$Searcher.Filter = "samaccountname=$($User.sAMAccountName)"
			$TmpUser = $Searcher.FindOne()
			# Use DirectoryEntry and RefreshCache method
			$CheckUser = $TmpUser.GetDirectoryEntry()
			$CheckUser.RefreshCache($Attribute)
			# Return the value of msds-user-account-control-computed
			$UserAccountFlag = $CheckUser.Properties[$Attribute].Value
			foreach ($key in $ADS_USER_FLAG_ENUM.Keys)
				{
				if ($UserAccountFlag -band $ADS_USER_FLAG_ENUM[$key])
					{
					$User.$key = $true
					}
				else
					{
					$User.$key = $false
					}
				}
			$FinalUsers +=$User
			}

		}
		catch
			{
				Write-Host "Error in Get-MsDSUACResolv function: $($_.Exception.Message)" -ForegroundColor Red
			}
	return $FinalUsers		
																				
	}



#---------------------------------------------------
# Resolv Bits of UserAccountControl
#---------------------------------------------------

function UACBitResolv ([Object] $Users)
	{
        Write-Host "Resolving UAC bits..." -ForegroundColor Green
		$FinalUsers = @()		
		foreach ($User in $Users)
			{
				foreach ($key in $UAC_USER_FLAG_ENUM.Keys)
					{
						if ([int][string]$User.userAccountControl -band $UAC_USER_FLAG_ENUM[$key])
							{
								$User.$key = $true
							}
						else
							{
								$User.$key = $false
							}
					}
				$FinalUsers += $User
			}
		return $FinalUsers
			
	}


#---------------------------------------------------
# Search for users
#---------------------------------------------------
function Get-UsersInfo ([String] $sAMAccountName, [Array] $Properties, [String] $Filter)
	{
        $Users = @()
	   # Filter on all user objects.
	   if($Global:Domain)
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
			$Searcher.Filter = "(&(objectCategory=person)(objectClass=user))"
			$Results = $Searcher.FindAll()
			Write-Verbose "The users search return $($Results.Count)" -Verbose
		}
	   else
		{
			$Searcher.filter = "(&(objectClass=user)(sAMAccountName= $($sAMAccountName)))"
			$Results = $Searcher.FindAll()
			#Write-Verbose "The user search return $($Results.Count)" -Verbose
		}
        ForEach ($Result In $Results)
		{
			$record = @{}
            ForEach ($Property in $Properties)
				{
                    if ($property -eq "pwdLastSet")
						{
                            if([string]::IsNullOrEmpty($Result.Properties.Item($property)))
                                {
                                    $record.Add($Property,"never")
                                }
                             else
                                {
                                    $pwlastSet = [datetime]::FromFileTime([Int64]::Parse($Result.Properties.Item($Property)))
							        $record.Add($Property,$pwlastSet)	
                                }                      							
						}
                     elseif ($Property -eq "accountExpires")
						{
							if([string]::IsNullOrEmpty($Result.Properties.Item($property)))
                                {
                                    $record.Add($Property,"never")                                    
                                }
                            elseif ($Result.Properties.Item($Property) -gt [datetime]::MaxValue.Ticks -or [Int64]::Parse($Result.Properties.Item($Property)) -eq 0)
								{
									$record.Add($Property,"Never")
								}
                            else
								{
									$Date = [Datetime]([Int64]::Parse($Result.Properties.Item($Property)))
									$accountExpires = $Date.AddYears(1600).ToLocalTime()
									$record.Add($Property,$accountExpires)
								}
						}
                      elseif ($Property -eq "lastLogonTimestamp")
						{
							$value = ($Result.Properties.Item($Property)).Count
                            switch($value)
										{
										  0
											{
											 $date = $Result.Properties.Item("WhenCreated")
											 $record.Add($Property,$null)
											}
										 default
											{
											$date = [datetime]::FromFileTime([Int64]::Parse($Result.Properties.Item($Property)))
											$record.Add($Property,[String]::Format($date))
											}
										}
                        }
                      else
						{
							$record.Add($Property,[String]$Result.Properties.Item([String]::Format($Property)))	
						}
                         
                    }
                $Users +=$record
        }
     return $Users
	}

<#
$tmp = Get-CurrentDomain
$Domain = $tmp
$Properties = @("distinguishedName","sAMAccountName","givenName","sn","description","userAccountControl"`
				,"pwdLastSet","whenCreated","whenChanged")
Get-UsersInfo -Domain $Domain.distinguishedName -Properties $Properties -uniq $false -sAMAccountName "silverhack"
#>

