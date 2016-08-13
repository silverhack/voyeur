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
            # Construct detailed PSObject
            #---------------------------------------------------

            function Set-PSObject ([Object] $Object, [String] $type){
		        $FinalObject = @()
		        foreach ($Obj in $Object){
				    $NewObject = New-Object -TypeName PSObject -Property $Obj
            	    $NewObject.PSObject.typenames.insert(0,[String]::Format($type))
				    $FinalObject +=$NewObject
			    }
		        return $FinalObject
	        }
            #----------------------
            #Start computers plugin
            #----------------------
            $GroupType = @{
			                "-2147483646" = "Global Security Group";
			                "-2147483644" = "Local Security Group";
			                "-2147483643" = "BuiltIn Group";
			                "-2147483640" = "Universal Security Group";
			                "2" = "Global Distribution Group";
			                "4" = "Local Distribution Group";
			                "8" = "Universal Distribution Group";}
            #----------------------
            #Start computers plugin
            #----------------------
            $DomainName = $ADObject.Domain.Name
            $PluginName = $ADObject.PluginName
            Write-Host ("{0}: {1}" -f "Groups task ID $bgRunspaceID", "Retrieve groups from $DomainName")`
            -ForegroundColor Magenta
            }
        Process{
            
            #Extract groups from domain
            $Domain = $ADObject.Domain
            $Connection = $ADObject.Domain.ADConnection
            $UseCredentials = $ADObject.UseCredentials
            $Filter = $ADObject.SearchRoot
            $Groups = @()
	        if($Connection){
		        $Searcher = New-Object System.DirectoryServices.DirectorySearcher -ArgumentList $Connection
			    $Searcher.SearchScope = "subtree"
                $Searcher.SearchRoot = $Connection
            }
            if($Searcher){
                $Searcher.PageSize = 200
            }
            # Add Attributes
            $GroupProperties = $ADObject.GroupFilter
            if($GroupProperties){
                foreach($property in $GroupProperties){
                    [void]$Searcher.PropertiesToLoad.Add([String]::Format($property))
                }
            }
            else{
                $GroupProperties = "*"
            }
            if (!$sAMAccountName){
		        $Searcher.Filter = "(&(objectCategory=group)(objectClass=group))"
			    $Results = $Searcher.FindAll()
			    Write-Verbose "Group search return $($Results.Count)" -Verbose
            }
            else{
			    $Searcher.filter = "(&(objectClass=group)(sAMAccountName= $sAMAccountName))"
			    $Results = $Searcher.FindAll()
            }
            if($Results){
                foreach ($result in $Results){
				    foreach ($Property in $GroupProperties){
                        $record = @{}
                        if($Property -eq "groupType"){
                            [String]$tmp = $Result.Properties.Item($Property)
                            $record.Add($Property,$GroupType[$tmp])	
                        }
                        else{
                            $record.Add($Property,[String]$Result.Properties.Item([String]::Format($Property)))
                        }
                    }
                    $Groups +=$record
			    }		
	        }
        }
        End{
            #Set PsObject from All Groups
            $AllGroups = Set-PSObject -Object $Groups -type "AD.Arsenal.Groups"
            #Work with SyncHash
            $SyncServer.$($PluginName)=$AllGroups

            #Create custom object for store data
            $GroupData = New-Object -TypeName PSCustomObject
            $GroupData | Add-Member -type NoteProperty -name Data -value $AllGroups

            #Formatting Excel Data
            $Excelformatting = New-Object -TypeName PSCustomObject
            $Excelformatting | Add-Member -type NoteProperty -name Data -value $AllGroups
            $Excelformatting | Add-Member -type NoteProperty -name TableName -value "All Groups"
            $Excelformatting | Add-Member -type NoteProperty -name SheetName -value "All Groups"
            $Excelformatting | Add-Member -type NoteProperty -name isFreeze -value $True
            $Excelformatting | Add-Member -type NoteProperty -name Type -value "CSVTable"

            #Add Excel formatting into psobject
            $GroupData | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting
            
            #Add Groups data to object
            $ReturnServerObject | Add-Member -type NoteProperty -name Groups -value $GroupData

            #Add Groups to report
            $CustomReportFields = $ReturnServerObject.Report
            $NewCustomReportFields = [array]$CustomReportFields+="AllGroups"
            $ReturnServerObject | Add-Member -type NoteProperty -name Report -value $NewCustomReportFields -Force
            #End
        }