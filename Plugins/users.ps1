#Plugin extract users from AD
[cmdletbinding()]
    Param (
            [Parameter(HelpMessage="Background Runspace ID")]
            [int]
            $bgRunspaceID,

            [Parameter(HelpMessage="Not used in this version")]
            [HashTable]
            $SyncServer,

            [Parameter(HelpMessage="Object with AD valuable data")]
            [Object]
            $ADObject,

            [Parameter(HelpMessage="Object to return data")]
            [Object]
            $ReturnServerObject,

            [Parameter(HelpMessage="Not used in this version")]
            [String]
            $sAMAccountName,

            [Parameter(HelpMessage="Organizational Unit filter")]
            [String]
            $Filter,

            [Parameter(HelpMessage="Attributes for user object")]
            [Array]
            $MyProperties = @("*")

        )
    Begin{

        #-----------------------------------------------------------
        # Function to get some stats about users
        #-----------------------------------------------------------
        Function Get-UsersStats{
            Param (
                [parameter(Mandatory=$true, HelpMessage="Object users")]
                [Object]$Status
            )

            # Dashboard users
			$UsersVolumetry = @{ 
								Actives = @($Status['isActive']['Count']); 
								Inactives = @($Status['isInActive']['Count'])
                                }

            $UsersStatus = @{ 
						Locked = @($Status['isActive']['isLocked'],$Status['isInActive']['isLocked']);
						"Password Expired" = @($Status['isActive']['isPwdExpired'],$Status['isInActive']['isPwdExpired']);
						Expired = @($Status['isActive']['isExpired'],$Status['isInActive']['isExpired']);
						Disabled = @($Status['isActive']['isDisabled'],$Status['isInActive']['isDisabled']);
						}

            $UsersConfig = @{ 
						"No delegated privilege" = @($Status['isActive']['isPTH'],$Status['isInActive']['isPTH']);
						"Password older than $($ADObject.InactiveDays) days" = @($Status['isActive']['isPwdOld'],$Status['isInActive']['isPwdOld']);
						"Password not required" = @($Status['isActive']['isPwdNotRequired'],$Status['isInActive']['isPwdNotRequired']);
						"Password never expires" = @($Status['isActive']['isPwdNeverExpires'],$Status['isInActive']['isPwdNeverExpires']);
					}
            #Return data
            return $UsersVolumetry, $UsersStatus, $UsersConfig

        }

        #-----------------------------------------------------------
        # Function to count High privileged users
        #-----------------------------------------------------------
        Function Get-PrivilegedStats{
            Param (
                [parameter(Mandatory=$true, HelpMessage="Object users")]
                [Object]$HighUsers
            )
            Begin{
                #Declare var
                $HighChart = @{}

            }
            Process{
                #Group for each object for count values
                $HighUsers | Group Group | ForEach-Object {$HighChart.Add($_.Name,@($_.Count))}           
            }
            End{
                if($HighChart){
                    return $HighChart
                }
            }
        }
        
        #-----------------------------------------------------------
        # Determine High Privileges in Active Directory object users
        #-----------------------------------------------------------

        function Get-HighPrivileges{
            Param (
                [parameter(Mandatory=$true, HelpMessage="Object users")]
                [Object]$Users
            )
            $Domain = $ADObject.Domain
            $Connection = $ADObject.Domain.ADConnection
            Write-Host "Trying to enumerate high privileges in $($Domain.Name)..." -ForegroundColor Magenta
            #Extract info from ADObject
            $DomainSID = $ADObject.DomainSID
            #End extraction data
            #http://support.microsoft.com/kb/243330
            $PrivGroups = @("S-1-5-32-552","S-1-5-32-544";"S-1-5-32-548";"S-1-5-32-549";"S-1-5-32-551";`
                        "$DomainSID-519";"$DomainSID-518";"$DomainSID-512";"$DomainSID-521")
            $HighPrivileges = @()
            #Create connection for search in AD
            $Searcher = New-Object System.DirectoryServices.DirectorySearcher -ArgumentList $Connection
            $Searcher.SearchScope = "subtree"
			$Searcher.SearchRoot = $Connection			
			$Searcher.PageSize = 200
            try{
                $ListPrivGroups = @{}
		        foreach ($group in $PrivGroups){
			        $Searcher.filter = "(objectSID=$group)"
        	        $Results = $Searcher.FindOne()
			        ForEach ($result in $Results){
            	        $dn = $result.properties.distinguishedname
				        $cn = $result.properties.cn
            	        $ListPrivGroups.Add($cn,$dn)
            	    }
			    }
                $ListPrivGroups.GetEnumerator() | ForEach-Object `
                    {
                        try{
                            $ErrorActionPreference = "Continue"
			                $Searcher.PageSize = 200
			                $Searcher.Filter = "(&(objectCategory=group)(name=$($_.Key)))"
			                $Items = $Searcher.FindAll()
                        }
                        catch{
                              $ErrorActionPreference = "Continue"
                              Write-Host "Function Get-HighPrivileges. Error in group $($_.Key)" -ForegroundColor Red
                        }
                        foreach ($Item in $Items){
					        $Group = $Item.GetDirectoryEntry()
					        $Members = $Group.member				
				        }
                        try{
                            if (($Group.Member).Count -gt 0){
                                foreach	($Member in $Members){
                                    $Properties = @("pwdLastSet","whenCreated")
						            $sAMAccountName = ($Member.replace('\','') -split ",*..=")[1]
                                    $Searcher.Filter = "(&(objectCategory=group)(name=$($_.Key)))"
                                    $UserFound = $Searcher.FindAll()
                                    $Match = $Users | Where-Object {$_.distinguishedName -eq $Member}
                                    if($Match.sAMAccountName){
                                        $HighPrivilege = @{
										"AccountName" = [String]::Format($sAMAccountName)
										"DN" = [String]::Format($Member)
										"MemberOf" = [String]::Format($_.Value)
										"PasswordLastSet" = $Match.pwdLastSet
										"whenCreated" = $Match.whenCreated
										"Group" = [String]::Format($_.Key)
										}
						                $obj = New-Object -TypeName PSObject -Property $HighPrivilege
            			                $obj.PSObject.typenames.insert(0,'Arsenal.AD.HighPrivileges')
            			                $HighPrivileges += $obj 
                                    }
			                                                           
                                }
                            }
                        }
                        catch{
                            Write-Warning ("{0}: {1}" -f "Error in function Get-HighPrivileges",$_.Exception.Message)
                        }
                    }
                    return $HighPrivileges
            }
            catch{
                Write-Warning ("{0}: {1}" -f "Error in function Get-HighPrivileges",$_.Exception.Message)
            }
        }

        #---------------------------------------------------
        # General Status of Active Directory object users
        #--------------------------------------------------- 

        function Get-UserStatus{
            Param (
                [parameter(Mandatory=$true, HelpMessage="Object users")]
                [Object]$Users
            )
            # Count user info
            $UsersCount = @{
				            isActive = @{
				            Count = 0;
				            isExpired = 0;
				            isDisabled = 0;
				            isPwdOld = 0;
				            isLocked = 0;
				            isPwdNeverExpires = 0;
				            isPwdNotRequired = 0;
				            isPwdExpired = 0
				            isPTH = 0;
				            }
				            isInactive = @{
				            Count = 0;
				            isExpired = 0;
				            isDisabled = 0;
				            isPwdOld = 0;
				            isLocked = 0;
				            isPwdNeverExpires = 0;
				            isPwdNotRequired = 0;
				            isPwdExpired = 0
				            isPTH = 0;
				            }
		            }
            $DN = $ADObject.Domain.GetDirectoryEntry().distinguishedName
		    for ($i=0; $i -lt $Users.Count; $i++){
			    Write-Progress -activity "Collecting Users metrics..." -status "$($DN): processing of $i/$($Users.Count) Users" -PercentComplete (($i / $Users.Count) * 100)
				# Active accounts
				if (([string]::IsNullOrEmpty($Users[$i].lastlogontimestamp)) -or (([DateTime]$Users[$i].lastlogontimestamp).ticks -lt $InactiveDate)){
				    $Users[$i].isActive = $false
				    $UsersCount['isInactive']['Count']++	
				}
				else{
					$Users[$i].isActive = $true
					$UsersCount['isActive']['Count']++
				}
				#Password Expires
				if ($Users[$i].ADS_UF_PASSWORD_EXPIRED -eq $true){
				    if ($Users[$i].isActive){
						$UsersCount['isActive']['isPwdExpired']++
					}
					else{
						$UsersCount['isInactive']['isPwdExpired']++
					}
				}
				if ($Users[$i].accountExpires -lt $CurrentDate){
				    $Users[$i].isExpired = $true
					if ($Users[$i].isActive){
					    $UsersCount['isActive']['isExpired']++
					}
					else{
						$UsersCount['isInactive']['isExpired']++
					}
				}
				else{
					$Users[$i].isExpired = $false						
				}
				if ($Users[$i].isDisabled){
					if ($Users[$i].isActive){
					    $UsersCount['isActive']['isDisabled']++
					}
					else{
						$UsersCount['isInactive']['isDisabled']++
					}
				}
				if ($Users[$i].isPwdNoRequired){
					if ($Users[$i].isActive){
						$UsersCount['isActive']['isPwdNotRequired']++
					}
					else{
						$UsersCount['isInactive']['isPwdNotRequired']++
					}
				}
				if ($Users[$i].isNoDelegation -eq $false){
					if ($Users[$i].isActive){
						$UsersCount['isActive']['isPTH']++
					}
					else{
						$UsersCount['isInactive']['isPTH']++
					}
				}
				# Identify pwdLastSet older than $PasswordAge days
				if ( $Users[$i].pwdLastSet -lt $PasswordLastSet ){
					$Users[$i].isPwdOld = $true
					if ( $UserObj.isActive ){
						$UsersCount['isActive']['isPwdOld']++
					}
					else{
						$UsersCount['isInactive']['isPwdOld']++
					}
				}
				# Identify password never expires
				if ( $Users[$i].isPwdNeverExpires){							
					if ( $Users[$i].isActive ){
						$UsersCount['isActive']['isPwdNeverExpires']++
					}
					else{
						$UsersCount['isInactive']['isPwdNeverExpires']++
					}
				}
			}
		    return $UsersCount
	    }
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
        # Resolv Bits of UserAccountControl
        #---------------------------------------------------

        function UACBitResolv{
             Param (
                [parameter(Mandatory=$true, HelpMessage="Users Object")]
                [Object]$Users
            )
            #Bits of UserAccountControl
            $UAC_USER_FLAG_ENUM = @{
						isLocked = 16;
						isDisabled = 2;
						isPwdNeverExpires = 65536;
						isNoDelegation = 1048576;
						isPwdNoRequired = 32;
						isPwdExpired = 8388608;
            }
            Write-Host "Resolving UAC bits..." -ForegroundColor Green
	        $FinalUsers = @()		
	        foreach ($User in $Users){
	            foreach ($key in $UAC_USER_FLAG_ENUM.Keys){
			        if ([int][string]$User.userAccountControl -band $UAC_USER_FLAG_ENUM[$key]){
			            $User.$key = $true
			        }
			        else{
				        $User.$key = $false}
			        }
		        $FinalUsers += $User
	        }
	        return $FinalUsers
        }
        #---------------------------------------------------
        # Resolv Bits of msds-user-account-control-computed
        #---------------------------------------------------

        function Get-MsDSUACResolv{
             Param (
                [parameter(Mandatory=$true, HelpMessage="Users object")]
                [Object]$Users
            )
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
            Write-Host "Resolving mds-UAC bits..." -ForegroundColor Green
            $FinalUsers = @()
            try{
                foreach ($User in $Users){
	                $Attribute = "msds-user-account-control-computed"
                    $CheckUser = $Result.GetDirectoryEntry()
                    $CheckUser.RefreshCache($Attribute)
                    $UserAccountFlag = $CheckUser.Properties[$Attribute].Value
                    foreach ($key in $ADS_USER_FLAG_ENUM.Keys){
				        if ($UserAccountFlag -band $ADS_USER_FLAG_ENUM[$key]){
					        $User.$key = $true
                        }
				        else{
					        $User.$key = $false
					    }
				    }
			        $FinalUsers +=$User
                }
                return $FinalUsers
            }
            catch{
                Write-Warning ("{0}: {1}" -f "Error in function Get-MsDSUACResolv",$_.Exception.Message)
            }
        }

        #---------------------------------------------------
        # Get user information from AD
        #---------------------------------------------------

        function Get-UsersInfo{
             Param (
                [parameter(Mandatory=$false, HelpMessage="DN or sAMAccountName to search")]
                [String]$AccountName,

                [parameter(Mandatory=$false, HelpMessage="Psobject with AD data")]
                [Object]$MyADObject,

                [parameter(Mandatory=$false, HelpMessage="Extract All properties")]
                [Switch]$ExtractAll
            )

            Begin{
                #Extract data from ADObject
                $Connection = $MyADObject.Domain.ADConnection
                $Filter = $MyADObject.SearchRoot
                $UsersProperties = $MyADObject.UsersFilter

                #Create Connection
                if($Connection){
		            $Searcher = New-Object System.DirectoryServices.DirectorySearcher -ArgumentList $Connection
			        $Searcher.SearchScope = "subtree"
                    $Searcher.SearchRoot = $Connection
                }
                #Add Pagesize property
                if($Searcher){
                    $Searcher.PageSize = 200
                }
                #Add properties to load for each user
                if(!$ExtractAll){
                    foreach($property in $UsersProperties){
                        $Searcher.PropertiesToLoad.Add([String]::Format($property)) > $Null
                    }
                }
                else{
                    $UsersProperties = "*"
                }
                if (!$sAMAccountName){
		            $Searcher.Filter = "(&(objectCategory=person)(objectClass=user))"
			        $Results = $Searcher.FindAll()
                }
                else{
			        $Searcher.filter = "(&(objectClass=user)(sAMAccountName= $($sAMAccountName)))"
			        $Results = $Searcher.FindAll()
                }

            }
            Process{
                if($Results){
                    Write-Verbose "The users search return $($Results.Count)" -Verbose
                    return $Results
                }
                
            }
            End{
                #Nothing to do here...                
            }

        }

        #Start users plugin
        $DomainName = $ADObject.Domain.Name
        $PluginName = $ADObject.PluginName
        Write-Host ("{0}: {1}" -f "Users task ID $bgRunspaceID", "Retrieve users data from $DomainName")`
        -ForegroundColor Magenta
    }
    Process{
        #Extract all users from domain
        <#
        $Connection = $ADObject.Domain.ADConnection
        $Filter = $ADObject.SearchRoot
        $UseCredentials = $ADObject.UseCredentials
        $Users = @()
	    if($Connection){
		    $Searcher = New-Object System.DirectoryServices.DirectorySearcher -ArgumentList $Connection
			$Searcher.SearchScope = "subtree"
            $Searcher.SearchRoot = $Connection
        }
        if($Searcher){
            $Searcher.PageSize = 200
        }
        # Add Attributes
        $UsersProperties = $ADObject.UsersFilter
        if($UsersProperties){
            foreach($property in $UsersProperties){
                $Searcher.PropertiesToLoad.Add([String]::Format($property)) > $Null
            }
        }
        else{
            $UsersProperties = "*"
        }
        if (!$sAMAccountName){
		    $Searcher.Filter = "(&(objectCategory=person)(objectClass=user))"
			$Results = $Searcher.FindAll()
			Write-Verbose "The users search return $($Results.Count)" -Verbose
        }
        else{
			$Searcher.filter = "(&(objectClass=user)(sAMAccountName= $($sAMAccountName)))"
			$Results = $Searcher.FindAll()
			#Write-Verbose "The user search return $($Results.Count)" -Verbose
        }
        #>
        $RawUsers = @()
        $Properties = $ADObject.UsersFilter
        $RawResults = Get-UsersInfo -MyADObject $ADObject
        if($RawResults){
            ForEach ($Result In $RawResults){
			    $record = @{}
                ForEach ($Property in $Properties){
                    if ($property -eq "pwdLastSet"){
                        if([string]::IsNullOrEmpty($Result.Properties.Item($property))){
                            $record.Add($Property,"never")
                        }
                        else{
                            $pwlastSet = [datetime]::FromFileTime([Int64]::Parse($Result.Properties.Item($Property)))
							$record.Add($Property,$pwlastSet)	
                        }                      							
					}
                    elseif ($Property -eq "accountExpires"){
					    if([string]::IsNullOrEmpty($Result.Properties.Item($property))){
                            $record.Add($Property,"never")                                    
                        }
                        elseif ($Result.Properties.Item($Property) -gt [datetime]::MaxValue.Ticks -or [Int64]::Parse($Result.Properties.Item($Property)) -eq 0){
						    $record.Add($Property,"Never")
						}
                        else{
							$Date = [Datetime]([Int64]::Parse($Result.Properties.Item($Property)))
							$accountExpires = $Date.AddYears(1600).ToLocalTime()
							$record.Add($Property,$accountExpires)
						}
					}
                    elseif ($Property -eq "lastLogonTimestamp"){
					    $value = ($Result.Properties.Item($Property)).Count
                        switch($value){
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
                    else{
						$record.Add($Property,[String]$Result.Properties.Item([String]::Format($Property)))
                        
					}
                         
                }
                $RawUsers +=$record
            }
        }
        
    }
    End{
        #Resolv UAC bits for each user
        $UsersWithUACResolv = UACBitResolv -Users $RawUsers

        #Resolv mds-UAC bits for each user
        $UsersWithmdsUACResolv = Get-MsDSUACResolv -Users $UsersWithUACResolv

        #Resolv High privileges in Domain
        $HighPrivileges = Get-HighPrivileges -Users $UsersWithmdsUACResolv

        #Resolv Status for each user
        $UsersStatus = Get-UserStatus -Users $UsersWithmdsUACResolv

        #Set PsObject from users object
        $FinalUsers = Set-PSObject -Object $UsersWithmdsUACResolv -type "AD.Arsenal.Users"
        
        #Work with SyncHash
        $SyncServer.$($PluginName)=$FinalUsers
        $SyncServer.UsersStatus=$UsersStatus

        ###########################################Reporting Options###################################################

        #Create custom object for store data
        $AllDomainUsers = New-Object -TypeName PSCustomObject
        $AllDomainUsers | Add-Member -type NoteProperty -name Data -value $FinalUsers

        #Formatting Excel Data
        $Excelformatting = New-Object -TypeName PSCustomObject
        $Excelformatting | Add-Member -type NoteProperty -name Data -value $FinalUsers
        $Excelformatting | Add-Member -type NoteProperty -name TableName -value "Domain Users"
        $Excelformatting | Add-Member -type NoteProperty -name SheetName -value "Domain Users"
        $Excelformatting | Add-Member -type NoteProperty -name isFreeze -value $True
        $Excelformatting | Add-Member -type NoteProperty -name Type -value "CSVTable"

        #Add Excel formatting into psobject
        $AllDomainUsers | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting
            
        #Add Users data to object
        if($FinalUsers){
            $ReturnServerObject | Add-Member -type NoteProperty -name DomainUsers -value $AllDomainUsers
        }

        #Add High privileges account to report
        $HighPrivilegedUsers = New-Object -TypeName PSCustomObject
        $HighPrivilegedUsers | Add-Member -type NoteProperty -name Data -value $HighPrivileges

        #Formatting Excel Data
        $Excelformatting = New-Object -TypeName PSCustomObject
        $Excelformatting | Add-Member -type NoteProperty -name Data -value $HighPrivileges
        $Excelformatting | Add-Member -type NoteProperty -name TableName -value "High Privileged Users"
        $Excelformatting | Add-Member -type NoteProperty -name SheetName -value "High Privileged Users"
        $Excelformatting | Add-Member -type NoteProperty -name isFreeze -value $True
        $Excelformatting | Add-Member -type NoteProperty -name Type -value "CSVTable"

        #Add Excel formatting into psobject
        $HighPrivilegedUsers | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting

        #Add Users data to object
        if($HighPrivileges){
            $ReturnServerObject | Add-Member -type NoteProperty -name HighPrivileged -value $HighPrivilegedUsers
        }


        #######################################Add Table options Users Status##########################################
        #Add new table for High privileges account to report
        $PrivilegedStats = Get-PrivilegedStats -HighUsers $HighPrivileges

        if($PrivilegedStats){
            $HighPrivilegedUsers = New-Object -TypeName PSCustomObject
            $HighPrivilegedUsers | Add-Member -type NoteProperty -name Data -value $PrivilegedStats

            #Formatting Excel Data
            $Excelformatting = New-Object -TypeName PSCustomObject
            $Excelformatting | Add-Member -type NoteProperty -name Data -value $PrivilegedStats
            $Excelformatting | Add-Member -type NoteProperty -name TableName -value "Privileged Users"
            $Excelformatting | Add-Member -type NoteProperty -name SheetName -value "Chart Privileged Users"
            $Excelformatting | Add-Member -type NoteProperty -name isnewSheet -value $True
            $Excelformatting | Add-Member -type NoteProperty -name showTotals -value $True
            $Excelformatting | Add-Member -type NoteProperty -name showHeaders -value $True
            $Excelformatting | Add-Member -type NoteProperty -name addHeader -value @('Type of group','Count')
            $Excelformatting | Add-Member -type NoteProperty -name position -value @(2,1)
            $Excelformatting | Add-Member -type NoteProperty -name Type -value "NewTable"

            #Add chart
            $Excelformatting | Add-Member -type NoteProperty -name addChart -value $True
            $Excelformatting | Add-Member -type NoteProperty -name chartType -value "xlColumnClustered"
            $Excelformatting | Add-Member -type NoteProperty -name ChartTitle -value "High Privileges Accounts in Groups"
            $Excelformatting | Add-Member -type NoteProperty -name style -value 34
            $Excelformatting | Add-Member -type NoteProperty -name hasDataTable -value $true

            $HighPrivilegedUsers | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting

            #Add Users status account
            $ReturnServerObject | Add-Member -type NoteProperty -name HighPrivilegedChart -value $HighPrivilegedUsers
         }
         
        
        #Add new table for Users account to report
        ($GlobalUserStats, $GlobalStatus, $GlobalUsersConfig)  = Get-UsersStats -Status $UsersStatus
        if($GlobalUserStats){
            $TmpCustomObject = New-Object -TypeName PSCustomObject
            #$GlobalStats | Add-Member -type NoteProperty -name Data -value $GlobalUserStats

            #Formatting Excel Data
            $Excelformatting = New-Object -TypeName PSCustomObject
            $Excelformatting | Add-Member -type NoteProperty -name Data -value $GlobalUserStats
            $Excelformatting | Add-Member -type NoteProperty -name TableName -value "Status of user Accounts"
            $Excelformatting | Add-Member -type NoteProperty -name SheetName -value "Users volumetry"
            $Excelformatting | Add-Member -type NoteProperty -name isnewSheet -value $true
            $Excelformatting | Add-Member -type NoteProperty -name showTotals -value $True
            $Excelformatting | Add-Member -type NoteProperty -name showHeaders -value $True
            $Excelformatting | Add-Member -type NoteProperty -name addHeader -value @('Type of account','Count')
            $Excelformatting | Add-Member -type NoteProperty -name position -value @(3,1)
            $Excelformatting | Add-Member -type NoteProperty -name Type -value "NewTable"

            #Add chart
            $Excelformatting | Add-Member -type NoteProperty -name addChart -value $True
            $Excelformatting | Add-Member -type NoteProperty -name chartType -value "xlPie"
            $Excelformatting | Add-Member -type NoteProperty -name ChartTitle -value "Volumetry of user accounts"
            $Excelformatting | Add-Member -type NoteProperty -name style -value 34
            $Excelformatting | Add-Member -type NoteProperty -name hasDataTable -value $true

            $TmpCustomObject | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting

            #Add Users status account
            $ReturnServerObject | Add-Member -type NoteProperty -name UsersVolumetry -value $TmpCustomObject
        }
        #Add status in the same sheet
        if($GlobalStatus){
            $TmpCustomObject = New-Object -TypeName PSCustomObject
            #$GlobalStats | Add-Member -type NoteProperty -name Data -value $GlobalUserStats

            #Formatting Excel Data
            $Excelformatting = New-Object -TypeName PSCustomObject
            $Excelformatting | Add-Member -type NoteProperty -name Data -value $GlobalStatus
            $Excelformatting | Add-Member -type NoteProperty -name TableName -value "User Account Status"
            $Excelformatting | Add-Member -type NoteProperty -name showTotals -value $True
            $Excelformatting | Add-Member -type NoteProperty -name showHeaders -value $True
            $Excelformatting | Add-Member -type NoteProperty -name addHeader -value @('Status','Active Accounts','Inactive Accounts')
            $Excelformatting | Add-Member -type NoteProperty -name isnewSheet -value $False
            $Excelformatting | Add-Member -type NoteProperty -name position -value @(3,6)
            $Excelformatting | Add-Member -type NoteProperty -name Type -value "NewTable"

            #Add chart
            $Excelformatting | Add-Member -type NoteProperty -name addChart -value $True
            $Excelformatting | Add-Member -type NoteProperty -name chartType -value "xlBarClustered"
            $Excelformatting | Add-Member -type NoteProperty -name ChartTitle -value "Status of active user accounts"
            $Excelformatting | Add-Member -type NoteProperty -name style -value 34
            $Excelformatting | Add-Member -type NoteProperty -name hasDataTable -value $true

            $TmpCustomObject | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting
            #Add Users status account
            $ReturnServerObject | Add-Member -type NoteProperty -name GlobalUsersStatus -value $TmpCustomObject
        }  
        
        if($GlobalUsersConfig){
            $TmpCustomObject = New-Object -TypeName PSCustomObject
            #$GlobalStats | Add-Member -type NoteProperty -name Data -value $GlobalUserStats

            #Formatting Excel Data
            $Excelformatting = New-Object -TypeName PSCustomObject
            $Excelformatting | Add-Member -type NoteProperty -name Data -value $GlobalUsersConfig
            $Excelformatting | Add-Member -type NoteProperty -name TableName -value "Configuration of user accounts"
            $Excelformatting | Add-Member -type NoteProperty -name isnewSheet -value $false
            $Excelformatting | Add-Member -type NoteProperty -name showTotals -value $True
            $Excelformatting | Add-Member -type NoteProperty -name showHeaders -value $True
            $Excelformatting | Add-Member -type NoteProperty -name addHeader -value @('Options','Active Accounts','Inactive Accounts')
            $Excelformatting | Add-Member -type NoteProperty -name position -value @(3,12)
            $Excelformatting | Add-Member -type NoteProperty -name Type -value "NewTable"

            #Add chart
            $Excelformatting | Add-Member -type NoteProperty -name addChart -value $True
            $Excelformatting | Add-Member -type NoteProperty -name chartType -value "xlColumnClustered"
            $Excelformatting | Add-Member -type NoteProperty -name ChartTitle -value "Configuration of active user accounts"
            $Excelformatting | Add-Member -type NoteProperty -name style -value 34
            $Excelformatting | Add-Member -type NoteProperty -name hasDataTable -value $true

            $TmpCustomObject | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting

            #Add Users status account
            $ReturnServerObject | Add-Member -type NoteProperty -name GlobalUsersConfig -value $TmpCustomObject
        }
        #Add users to report
        $CustomReportFields = $ReturnServerObject.Report
        $NewCustomReportFields = [array]$CustomReportFields+="Users, HighPrivileges, UserStatus"
        $ReturnServerObject | Add-Member -type NoteProperty -name Report -value $NewCustomReportFields -Force
        #End
    }