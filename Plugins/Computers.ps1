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
            # General Status of Active Directory objects
            #--------------------------------------------------- 

            Function Get-ComputerStatus{
                Param (
                [parameter(Mandatory=$true, HelpMessage="Computers Object")]
                [Object]$Computers
                )
                #Populate Array with OS & Service Pack
                $AllOS= @()
                for ($i=0; $i -lt $Computers.Count; $i++){
                    $OS = "{0} {1}" -f $Computers[$i].operatingSystem, $Computers[$i].operatingSystemServicePack
                    $NewOS = @{"OS"=$OS}
                    $obj = New-Object -TypeName PSObject -Property $NewOS
            		$obj.PSObject.typenames.insert(0,'Arsenal.AD.OperatingSystem')
                    $AllOS+=$obj
                }
                $OSChart = @{}
                $AllOS | Group OS | ForEach-Object {$OSChart.Add($_.Name,@($_.Count))}
                if($OSChart){
                    return $OSChart
                }
                
            }
        
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
            #Start computers plugin
            $DomainName = $ADObject.Domain.Name
            $PluginName = $ADObject.PluginName
            Write-Host ("{0}: {1}" -f "Computers task ID $bgRunspaceID", "Retrieve computers data from $DomainName")`
            -ForegroundColor Magenta
            }
        Process{
            #Extract computers from domain
            $Domain = $ADObject.Domain
            $Connection = $ADObject.Domain.ADConnection
            $UseSSL = $ADObject.UseSSL         
            $Computers = @()
	        if($Connection){
		        $Searcher = New-Object System.DirectoryServices.DirectorySearcher -ArgumentList $Connection
                $Searcher.SearchScope = "subtree"
                $Searcher.SearchRoot = $Connection
            }
            if($Searcher){
                $Searcher.PageSize = 200
            }
            # Add Attributes
            $ComputersProperties = $ADObject.ComputersFilter
            if($ComputersProperties){
                foreach($property in $ComputersProperties){
                    $Searcher.PropertiesToLoad.Add([String]::Format($property)) > $Null
                }
            }
            else{
                $ComputersProperties = "*"
            }
            if (!$sAMAccountName){
		        $Searcher.Filter = "(objectCategory=Computer)"
			    $Results = $Searcher.FindAll()
			    Write-Verbose "Computers search return $($Results.Count)" -Verbose
            }
            else{
			    $Searcher.filter = "(&(objectClass=Computer)(sAMAccountName= $sAMAccountName))"
			    $Results = $Searcher.FindAll()
			    #Write-Verbose "The user search return $($Results.Count)" -Verbose
            }
            if($Results){
                 ForEach ($Result In $Results){
			        $record = @{}
			        ForEach ($Property in $ComputersProperties){
                        if ($Property -eq "lastLogon"){
                            if ([String]::IsNullOrEmpty($Result.Properties.Item([String]::Format($Property)))){
                                $record.Add($Property,"Never")
                            }
                            else{
							    $lastLogon = [datetime]::FromFileTime([Int64]::Parse($Result.Properties.Item([String]::Format($Property))))
							    $record.Add($Property,$lastLogon)
                            }
						}
                        elseif ($Property -eq "pwdLastSet"){
                            if ([String]::IsNullOrEmpty($Result.Properties.Item([String]::Format($Property)))){
                                $record.Add($Property,0)
                            }
                            else{
							    $pwdLastSet = [datetime]::FromFileTime([Int64]::Parse($Result.Properties.Item([String]::Format($Property))))
							    $record.Add($Property,$pwdLastSet)
                            }
						}
                        else{
					        $record.Add($Property,[String] $Result.Properties.Item([String]::Format($Property)))	
                        }
				    }
			        $Computers +=$record
		        }
            }
        }
        End{
            #Get status from all computers
            $ComputersStatus = Get-ComputerStatus -Computers $Computers
            #$OSCount = $ComputersStatus | Format-Table @{label='OS';expression={$_.Name}},@{label='Count';expression={$_.Value}}
            #Set PsObject from computers object
            $AllComputers = Set-PSObject -Object $Computers -type "AD.Arsenal.Computers"
            #Work with SyncHash
            $SyncServer.$($PluginName)=$AllComputers
            $SyncServer.ComputersStatus=$ComputersStatus
            #Add computers data to object
            #$ReturnServerObject | Add-Member -type NoteProperty -name ComputersStatus -value $ComputersStatus

            #Create custom object for store data
            $ComputersData = New-Object -TypeName PSCustomObject
            $ComputersData | Add-Member -type NoteProperty -name Data -value $AllComputers

            #Formatting Excel Data
            $Excelformatting = New-Object -TypeName PSCustomObject
            $Excelformatting | Add-Member -type NoteProperty -name Data -value $AllComputers
            $Excelformatting | Add-Member -type NoteProperty -name TableName -value "All Computers"
            $Excelformatting | Add-Member -type NoteProperty -name SheetName -value "All Computers"
            $Excelformatting | Add-Member -type NoteProperty -name isFreeze -value $True
            $Excelformatting | Add-Member -type NoteProperty -name Type -value "CSVTable"

            #Add Excel formatting into psobject
            $ComputersData | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting

            #Add Groups data to object
            $ReturnServerObject | Add-Member -type NoteProperty -name Computers -value $ComputersData

            #Add Computer status Chart
            if($ComputersStatus){
                $TmpCustomObject = New-Object -TypeName PSCustomObject

                #Formatting Excel Data
                $Excelformatting = New-Object -TypeName PSCustomObject
                $Excelformatting | Add-Member -type NoteProperty -name Data -value $ComputersStatus
                $Excelformatting | Add-Member -type NoteProperty -name TableName -value "Computer status"
                $Excelformatting | Add-Member -type NoteProperty -name SheetName -value "Dashboard Computers"
                $Excelformatting | Add-Member -type NoteProperty -name isnewSheet -value $True
                $Excelformatting | Add-Member -type NoteProperty -name showTotals -value $True
                $Excelformatting | Add-Member -type NoteProperty -name showHeaders -value $false
                $Excelformatting | Add-Member -type NoteProperty -name addHeader -value $false
                $Excelformatting | Add-Member -type NoteProperty -name position -value @(2,1)
                $Excelformatting | Add-Member -type NoteProperty -name Type -value "NewTable"
                
                #Add chart
                $Excelformatting | Add-Member -type NoteProperty -name addChart -value $True
                $Excelformatting | Add-Member -type NoteProperty -name chartType -value "xlColumnClustered"
                $Excelformatting | Add-Member -type NoteProperty -name ChartTitle -value "OS versions"
                $Excelformatting | Add-Member -type NoteProperty -name style -value 34
                $Excelformatting | Add-Member -type NoteProperty -name hasDataTable -value $true
                
                $TmpCustomObject | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting
              
                #Add Computers chart
                $ReturnServerObject | Add-Member -type NoteProperty -name GlobalComputers -value $TmpCustomObject
                
            }


            #Add Computers to report
            $CustomReportFields = $ReturnServerObject.Report
            $NewCustomReportFields = [array]$CustomReportFields+="Computers, ComputerStatus"
            $ReturnServerObject | Add-Member -type NoteProperty -name Report -value $NewCustomReportFields -Force
            #End
        }