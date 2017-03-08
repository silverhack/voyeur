#Plugin extract roles from AD
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

            #-------------------------------------------------------
            # Function to convert msds-replvaluemetadata in XML data
            #-------------------------------------------------------
            function ConvertReplTo-XML{
                Param (
                    [parameter(Mandatory=$true, HelpMessage="msds-replvaluemetadata object")]
                    [Object]$replmetadata
                )
                Begin{
                    #Initialize vars
                    $NewXmlData = $Null                    
                }
                Process{
                    if($replmetadata -ne $Null){
                        $NewXmlData = $replmetadata.Replace("&","&amp;")
                    }
                }
                End{
                    if($NewXmlData){
                        return [XML]$NewXmlData
                    } 
                }
            }

            #---------------------------------------------------
            # Function to analyze msds-replvaluemetadata info
            #---------------------------------------------------
            function Get-MetadataInfo{
                Param (
                    [parameter(Mandatory=$true, HelpMessage="Groups object")]
                    [Object]$AllGroups
                )
                Begin{
                    $Threshold = 180
                    $Today = Get-Date
                    $AllObj = @()
                }
                Process{
                    for ($i=0; $i -lt $AllGroups.Count; $i++){
                        $distinguishedname = [string] $AllGroups[$i].Properties.distinguishedname
                        $sAMAccountName = [string] $AllGroups[$i].Properties.samaccountname
                        $replmetatada = $AllGroups[$i].Properties.'msds-replvaluemetadata'
                        # Show progress in the script
                        Write-Progress -Activity "Collecting protected groups" `
                        -Status "$($sAMAccountName): processing of $i/$($AllGroups.Count) Groups"`
                        -Id 0 -CurrentOperation "Parsing group: $distinguishedname"`
                        -PercentComplete ( ($i / $AllGroups.Count) * 100 )
                        #Iterate for each object
                        foreach($metadata in $replmetatada){
                            if($metadata){
                                [XML]$XML = ConvertReplTo-XML -replmetadata $metadata
                                $Match= $xml.DS_REPL_VALUE_META_DATA | Where-Object {$_.pszAttributeName -eq "member"} |
                                Select-Object @{name="Attribute"; expression={$_.pszAttributeName}},
                                                @{name="objectDN";expression={$_.pszObjectDn}},
                                                @{name="pszObjectDn";expression={$_.ftimeCreated}},
                                                @{name="dwVersion";expression={$_.dwVersion}},
                                                @{name="ftimeDeleted";expression={$_.ftimeDeleted}},
                                                @{name="ftimeCreated";expression={$_.ftimeCreated}}
                                try{
                                    #Check if object has been added or removed
                                    If ( ([datetime] $Match.ftimeDeleted) -ge $Today.AddDays(-$Threshold)`
                                                 -or ([datetime] $Match.ftimeCreated) -ge $Today.AddDays(-$Threshold)){
                                        #Check if added or removed
                                        If ( $Match.ftimeDeleted -ne "1601-01-01T00:00:00Z"){
                                            #Deleted user
                                            $Flag = "removed"
                                            $Datemodified   = $Match.ftimeDeleted
                                        }
                                        else{
                                            #User added
                                            $Flag = "added"
                                            $Datemodified   = $Match.ftimeCreated
                                        }
                                        #Create Object
                                        $record = @{}
						                $record.Add("distinguishedName", $distinguishedname)
                                        $record.Add("sAMAccountName", $sAMAccountName)
                                        $record.Add("MemberDN", $Match.ObjectDn)
                                        $record.Add("Operation", $Flag)
                                        $record.Add("FirstTime", $Match.ftimeCreated)
                                        $record.Add("ModifiedTime", $Datemodified)
                                        $record.Add("Version", $Match.dwVersion)

                                        #Add to array
                                        $AllObj += $record
                                    }
                                }
                                catch{
                                    #Nothing to do here
                                }
                            }
                        }
                        
                    }
                }
                End{
                    if($AllObj){
                        return $AllObj
                    }
                }
            }

            #---------------------------------------------------
            # Construct detailed PSObject
            #---------------------------------------------------

            function Set-PSObject{
                Param (
                    [parameter(Mandatory=$true, HelpMessage="Object")]
                    [Object]$Object,

                    [parameter(Mandatory=$true, HelpMessage="Object")]
                    [String]$Type
                )
                Begin{
                    #Declare vars
                    $FinalObject = @()
                }
                Process{
                    foreach ($Obj in $Object){
				        $NewObject = New-Object -TypeName PSObject -Property $Obj
            	        $NewObject.PSObject.typenames.insert(0,[String]::Format($type))
				        $FinalObject +=$NewObject
			        }
                }
                End{
                    if($FinalObject){
                        return $FinalObject
                    }
                }
            }

            #Start computers plugin
            $DomainName = $ADObject.Domain.Name
            $PluginName = $ADObject.PluginName
            Write-Host ("{0}: {1}" -f "Group membership task ID $bgRunspaceID", "Retrieve group membership change information from $DomainName")`
            -ForegroundColor Magenta    
        }
        Process{
            #Extract information about connection
            $Connection = $ADObject.Domain.ADConnection
            if($Connection){
		        $Searcher = New-Object System.DirectoryServices.DirectorySearcher -ArgumentList $Connection
			    $Searcher.SearchScope = "subtree"
                $Searcher.SearchRoot = $Connection
            }
            if($Searcher){
                $Searcher.PageSize = 200
                #Add filter
                $Searcher.Filter = "(&(objectClass=group)(adminCount=1))"
                #Add properties to query
                [void]$Searcher.PropertiesToLoad.Add("distinguishedName")
                [void]$Searcher.PropertiesToLoad.Add("msDS-ReplValueMetaData")
                [void]$Searcher.PropertiesToLoad.Add("sAMAccountName")
                $PriviledgeGroups = $Searcher.FindAll()
                if($PriviledgeGroups){                    
                    $RawObj = Get-MetadataInfo -AllGroups $PriviledgeGroups                    
                }
            }
        }
        End{
            if($RawObj){
                $MonitoredGroups = Set-PSObject $RawObj "Arsenal.AD.MonitoredGroups"
                #Work with SyncHash
                $SyncServer.$($PluginName)=$MonitoredGroups

                #Create custom object for store data
                $TmpObject = New-Object -TypeName PSCustomObject
                $TmpObject | Add-Member -type NoteProperty -name Data -value $MonitoredGroups

                #Formatting Excel Data
                $Excelformatting = New-Object -TypeName PSCustomObject
                $Excelformatting | Add-Member -type NoteProperty -name Data -value $MonitoredGroups
                $Excelformatting | Add-Member -type NoteProperty -name TableName -value "Administrative Group changes"
                $Excelformatting | Add-Member -type NoteProperty -name SheetName -value "Administrative Group changes"
                $Excelformatting | Add-Member -type NoteProperty -name isFreeze -value $True
                $Excelformatting | Add-Member -type NoteProperty -name Type -value "CSVTable"

                #Add Excel formatting into psobject
                $TmpObject | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting

                #Add roles data to object
                $ReturnServerObject | Add-Member -type NoteProperty -name GroupChange -value $TmpObject
            }
        }