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

            #-----------------------------------------------------------
            # Function to count Services
            #-----------------------------------------------------------
            Function Get-ServiceStats{
                Param (
                    [parameter(Mandatory=$true, HelpMessage="Object")]
                    [Object]$AllServices
                )
                Begin{
                    #Declare var
                    $RolesChart = @{}
                }
                Process{
                    #Group for each object for count values
                    $AllServices | Group Service | ForEach-Object {$RolesChart.Add($_.Name,@($_.Count))}           
                }
                End{
                    if($RolesChart){
                        return $RolesChart
                    }
                }
            }
            #Common Service Principal Names
            $Services = @(@{Service="SQLServer";SPN="MSSQLSvc"},@{Service="TerminalServer";SPN="TERMSRV"},
					@{Service="Exchange";SPN="IMAP4"},@{Service="Exchange";SPN="IMAP"},
					@{Service="Exchange";SPN="SMTPSVC"},@{Service="SCOM";SPN="AdtServer"},
					@{Service="SCOM";SPN="MSOMHSvc"},@{Service="Cluster";SPN="MSServerCluster"},
					@{Service="Cluster";SPN="MSServerClusterMgmtAPI"};@{Service="GlobalCatalog";SPN="GC"},
					@{Service="DNS";SPN="DNS"},@{Service="Exchange";SPN="exchangeAB"}, 
                    @{Service ="WebServer";SPN="tapinego"},@{Service ="WinRemoteAdministration";SPN="WSMAN"},
                    @{Service ="ADAM";SPN="E3514235-4B06-11D1-AB04-00C04FC2DCD2-ADAM"},
                    @{Service ="Exchange";SPN="exchangeMDB"}
            )
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
            Write-Host ("{0}: {1}" -f "Roles task ID $bgRunspaceID", "Retrieve role information from $DomainName")`
            -ForegroundColor Magenta
            }
        Process{
            #Extract roles from domain
            $Domain = $ADObject.Domain
            $UseCredentials = $ADObject.UseCredentials
            $Connection = $ADObject.Domain.ADConnection
            $FinalObject = @()
	        if($Connection){
		        $Searcher = New-Object System.DirectoryServices.DirectorySearcher -ArgumentList $Connection
			    $Searcher.SearchScope = "subtree"
                $Searcher.SearchRoot = $Connection
            }
            if($Searcher){
                $Searcher.PageSize = 200
            }
            #Search roles in AD
            Foreach ($service in $Services){
			    $spn = [String]::Format($service.SPN)
			    $Searcher.filter = "(servicePrincipalName=$spn*/*)"
        	    $Results = $Searcher.FindAll()
			    if ($Results){
					Write-Verbose "Role Service Found....."
					foreach ($r in $Results){
						$account = $r.GetDirectoryEntry()
						$record = @{}
						$record.Add("Name", [String]$account.Name)
						$record.Add("AccountName", [String]$account.sAMAccountName)
						$record.Add("GUID", [String]$account.guid)
						$record.Add("DN", [String]$account.DistinguishedName)
						$record.Add("Service", [String]$service.Service)
						$FinalObject +=$record
					}
				}
		    }
        }
        End{
            if($FinalObject){
                #Set PsObject from roles found
                $AllRoles = Set-PSObject $FinalObject "Arsenal.AD.Roles"
                #Work with SyncHash
                $SyncServer.$($PluginName)=$AllRoles
            
                #Create custom object for store data
                $RoleData = New-Object -TypeName PSCustomObject
                $RoleData | Add-Member -type NoteProperty -name Data -value $AllRoles

                #Formatting Excel Data
                $Excelformatting = New-Object -TypeName PSCustomObject
                $Excelformatting | Add-Member -type NoteProperty -name Data -value $AllRoles
                $Excelformatting | Add-Member -type NoteProperty -name TableName -value "All roles"
                $Excelformatting | Add-Member -type NoteProperty -name SheetName -value "All Roles"
                $Excelformatting | Add-Member -type NoteProperty -name isFreeze -value $True
                $Excelformatting | Add-Member -type NoteProperty -name Type -value "CSVTable"

                #Add Excel formatting into psobject
                $RoleData | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting

                #Add roles data to object
                $ReturnServerObject | Add-Member -type NoteProperty -name Roles -value $RoleData

                #Add Chart for roles
                $RoleChart = Get-ServiceStats -AllServices $AllRoles
                if($RoleChart){
                    $TmpCustomObject = New-Object -TypeName PSCustomObject

                    #Formatting Excel Data
                    $Excelformatting = New-Object -TypeName PSCustomObject
                    $Excelformatting | Add-Member -type NoteProperty -name Data -value $RoleChart
                    $Excelformatting | Add-Member -type NoteProperty -name TableName -value "Service status"
                    $Excelformatting | Add-Member -type NoteProperty -name SheetName -value "Role Dashboard"
                    $Excelformatting | Add-Member -type NoteProperty -name isnewSheet -value $True
                    $Excelformatting | Add-Member -type NoteProperty -name showTotals -value $True
                    $Excelformatting | Add-Member -type NoteProperty -name showHeaders -value $false
                    $Excelformatting | Add-Member -type NoteProperty -name addHeader -value @('Type of role','Count')
                    $Excelformatting | Add-Member -type NoteProperty -name position -value @(2,1)
                    $Excelformatting | Add-Member -type NoteProperty -name Type -value "NewTable"

                    #Add chart
                    $Excelformatting | Add-Member -type NoteProperty -name addChart -value $True
                    $Excelformatting | Add-Member -type NoteProperty -name chartType -value "xlColumnClustered"
                    $Excelformatting | Add-Member -type NoteProperty -name ChartTitle -value "Role status"
                    $Excelformatting | Add-Member -type NoteProperty -name style -value 34
                    $Excelformatting | Add-Member -type NoteProperty -name hasDataTable -value $true
                
                    $TmpCustomObject | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting
              
                    #Add Computers chart
                    $ReturnServerObject | Add-Member -type NoteProperty -name GlobalServices -value $TmpCustomObject
                }
            
                #Add Computers to report
                $CustomReportFields = $ReturnServerObject.Report
                $NewCustomReportFields = [array]$CustomReportFields+="AllRoles"
                $ReturnServerObject | Add-Member -type NoteProperty -name Report -value $NewCustomReportFields -Force
                #End
            }
        }