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

#Plugin extract computers from AD
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

            #---------------------------------------------------
            # Get Icon of ACL
            #---------------------------------------------------

            Function Get-Icon{
                Param (
                [parameter(Mandatory=$true, HelpMessage="Color")]
                [String]$Color
                )
                #A few colors for format cells
                $ACLIcons = @{
			        "CreateChild, DeleteChild" = 66; #Warning
                    "DeleteTree, WriteDacl" = 66;#Warning
                    "Delete" = 99#Warning
                    "CreateChild, DeleteChild, ListChildren" = 66; #Warning
			        "CreateChild, Self, WriteProperty, ExtendedRight, Delete, GenericRead, WriteDacl, WriteOwner" = 66;#Warning
                    "CreateChild, DeleteChild, Self, WriteProperty, ExtendedRight, Delete, WriteDacl, WriteOwner" = 66;#Warning
			        "CreateChild, Self, WriteProperty, ExtendedRight, GenericRead, WriteDacl, WriteOwner" = 66;#Warning
                    "ListChildren, ReadProperty, ListObject" = 10; #Low
			        "GenericAll" = 99;#High
			        "GenericRead" = 10; #Low
			        "ReadProperty, WriteProperty, ExtendedRight" = 66;#Warning
			        "ListChildren" = 10;#Low
			        "ReadProperty" = 10;#Low
			        "ReadProperty, WriteProperty" = 66;#Warning
                    "ReadProperty, GenericExecute" = 66;#Warning
                    "GenericRead, WriteDacl" = 10; #Low
			        "ExtendedRight" = 66;#Warning
                    "WriteProperty" = 66;#Warning
                    "DeleteChild" = 66;#Warning
			        "CreateChild, DeleteChild, Self, WriteProperty, ExtendedRight, Delete, GenericRead, WriteDacl, WriteOwner" = 66;#Warning
			        "CreateChild, DeleteChild, Self, WriteProperty, ExtendedRight, GenericRead, WriteDacl, WriteOwner" = 66;#Warning
			        "DeleteTree, Delete" = 66;#Warning
                    "DeleteChild, DeleteTree, Delete" = 66;#Warning
                    "CreateChild, ReadProperty, GenericExecute" = 66; #Warning
                    "Self, WriteProperty, GenericRead" = 66; #Warning
                    "CreateChild, ReadProperty, WriteProperty, GenericExecute" = 66; #Warning
                    "ListChildren, ReadProperty, GenericWrite" = 66;#Warning
                    "CreateChild, ListChildren, ReadProperty, GenericWrite" = 66;#Warning
                    "CreateChild" = 10;#Low
			    }
		        try{
				    return $ACLIcons[$Color]
			    }
		        catch{
				      Write-Host "ACL color not found" -ForegroundColor Red
			    }
	        }
            
            #---------------------------------------------------
            # Search ACL in domain
            #---------------------------------------------------
            function TranslateInheritanceType{
                Param (
                [parameter(Mandatory=$true, HelpMessage="InheritanceType")]
                [String]$Type
                )
                $AppliesTo = ""
		        $InheritanceType = @{"None"="This object only";"Descendents"="All child objects";`
							 "SelfAndChildren"="This object and one level Of child objects";`
							 "Children"="One level of child objects";"All"="This object and all child objects"}
                try{
				    $AppliesTo = $InheritanceType[[String]::Format($type)]
			    }
		        catch{
				      $AppliesTo = "Flag not resolved"
				      Write-Debug $AppliesTo
			    }
		        return $AppliesTo
            }
            #---------------------------------------------------
            # Translate SID
            #---------------------------------------------------
            function TranslateSid{
                Param (
                [parameter(Mandatory=$true, HelpMessage="SID object to translate")]
                [Object]$SID
                )
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
                try{			
				    $Account = $SID.Translate([System.Security.Principal.NTAccount])
				    #Write-Host $Account -ForegroundColor Red
                }
		        catch{
				     $Account = $KnowSID[[String]::Format($SID)]	
				     #Write-Host $Account -ForegroundColor Yellow
			    }
		        return $Account
            }
        }
        Process{
            #Extract all users from domain
            $Domain = $ADObject.Domain
            $Connection = $ADObject.Domain.ADConnection
            $UseCredentials = $ADObject.UseCredentials
            $PluginName = $ADObject.PluginName
            #Extract action from object
            $Action = $ADObject.OUExtract
            $FullACLExtract = $Action.FullACL
            $Query = $Action.Query
            $ReportName = $Action.Name
            #End action retrieve
	        if($Connection){
		        $Searcher = New-Object System.DirectoryServices.DirectorySearcher -ArgumentList $Connection
			    $Searcher.SearchScope = "subtree"
                $Searcher.SearchRoot = $Connection
            }
            if($Searcher){
                $Searcher.PageSize = 200
            }
            if($Action){
                Write-Host ("{0}: {1}" -f "ACL task ID $bgRunspaceID", "Retrieve $($Action.Name) ACL and data from $($Domain.Name)")`
                -ForegroundColor Magenta
                $Searcher.filter = $Query
		        [void]$Searcher.PropertiesToLoad.AddRange("DistinguishedName")
		        [void]$Searcher.PropertiesToLoad.AddRange("name")
                $Results = $Searcher.FindAll()
                $ListOU = @()
                if($FullACLExtract -and $Results){
				    foreach ($result in $Results){
					    #$ListOU += $result.properties.item("DistinguishedName")
					    $data = $result.GetDirectoryEntry()
					    $aclObject = $data.Get_ObjectSecurity()
					    $aclList = $aclObject.GetAccessRules($true,$true,[System.Security.Principal.SecurityIdentifier])
					    foreach($acl in $aclList){
             			    $objSID = New-Object System.Security.Principal.SecurityIdentifier($acl.IdentityReference)
			 			    $AccountName = TranslateSID($objSID)
			 			    $AppliesTo = TranslateInheritanceType($acl.InheritanceType)
             			    $OUACL = @{
			 						    "AccountName" = [String]$AccountName
									    "DN" = [String]::Format($result.properties.item('DistinguishedName'))
									    "Rights" = $acl.ActiveDirectoryRights
									    "Risk" = (Get-Icon $acl.ActiveDirectoryRights).ToString()
									    "AppliesTo" = [String]$AppliesTo
									    "AccessControlType" = [String]$acl.AccessControlType
									    "InheritanceFlags" = [String]$acl.InheritanceFlags
						    } 
            			    $obj = New-Object -TypeName PSObject -Property $OUACL
            			    $obj.PSObject.typenames.insert(0,'Arsenal.AD.$($ReportName)')
						    $ListOU +=$obj
            			    #Write-Output $obj		
					    }
				    }
                    #Add Full OU data to object
                    #Create custom object for store data
                    $TmpObject = New-Object -TypeName PSCustomObject
                    $TmpObject | Add-Member -type NoteProperty -name Data -value $ListOU

                    #Formatting Excel Data
                    $Excelformatting = New-Object -TypeName PSCustomObject
                    $Excelformatting | Add-Member -type NoteProperty -name Data -value $ListOU
                    $Excelformatting | Add-Member -type NoteProperty -name TableName -value $ReportName
                    $Excelformatting | Add-Member -type NoteProperty -name SheetName -value $ReportName
                    $Excelformatting | Add-Member -type NoteProperty -name isFreeze -value $True
                    $Excelformatting | Add-Member -type NoteProperty -name IconSet -value $True
                    $Excelformatting | Add-Member -type NoteProperty -name IconColumnName -value "Risk"
                    $Excelformatting | Add-Member -type NoteProperty -name Type -value "CSVTable"

                    #Add Excel formatting into psobject
                    $TmpObject | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting

                    #Add ACL data to object
                    if($ListOU){
                        $ReturnServerObject | Add-Member -type NoteProperty -name $ReportName -value $TmpObject
                    }

                    #Add Full OU to report
                    $CustomReportFields = $ReturnServerObject.Report
                    $NewCustomReportFields = [array]$CustomReportFields+=$ReportName
                    $ReturnServerObject | Add-Member -type NoteProperty -name Report -value $NewCustomReportFields -Force
                    #Work with Synchronized hash
                    $SyncServer.FullOrganizationalUnit=$ListOU
                }
                if(!$FullACLExtract -and $Results){
				    foreach ($result in $Results){
					    $OUProperties = @{
									        "Name" = [String]::Format($result.properties.item('name'))
									        "DN" = [String]::Format($result.properties.item('DistinguishedName'))
						}
						$obj = New-Object -TypeName PSObject -Property $OUProperties
            			$obj.PSObject.typenames.insert(0,'Arsenal.AD.$($ReportName)')
						$ListOU +=$obj
					}
                    #Add Simple OU data to object
                    #Create custom object for store data
                    $TmpObject = New-Object -TypeName PSCustomObject
                    $TmpObject | Add-Member -type NoteProperty -name Data -value $ListOU

                    #Formatting Excel Data
                    $Excelformatting = New-Object -TypeName PSCustomObject
                    $Excelformatting | Add-Member -type NoteProperty -name Data -value $ListOU
                    $Excelformatting | Add-Member -type NoteProperty -name TableName -value $ReportName
                    $Excelformatting | Add-Member -type NoteProperty -name SheetName -value $ReportName
                    $Excelformatting | Add-Member -type NoteProperty -name isFreeze -value $True
                    #$Excelformatting | Add-Member -type NoteProperty -name IconSet -value $True
                    #$Excelformatting | Add-Member -type NoteProperty -name IconColumnName -value "Risk"
                    $Excelformatting | Add-Member -type NoteProperty -name Type -value "CSVTable"

                    #Add Excel formatting into psobject
                    $TmpObject | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting

                    #Add ACL data to object
                    if($ListOU){
                        $ReturnServerObject | Add-Member -type NoteProperty -name $ReportName -value $TmpObject
                    }	


                    #Add SimpleOU to report
                    $CustomReportFields = $ReturnServerObject.Report
                    $NewCustomReportFields = [array]$CustomReportFields+="$ReportName"
                    $ReturnServerObject | Add-Member -type NoteProperty -name Report -value $NewCustomReportFields -Force
                    #Work with Synchronized hash
                    $SyncServer.OrganizationalUnit=$ListOU	
                }
            }            
            
        
        }
        End{
            #Nothing to do here
            #End
        }