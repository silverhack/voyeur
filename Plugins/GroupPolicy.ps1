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
            $AllGPO = @()
            $DomainName = $ADObject.Domain.Name
            $PluginName = $ADObject.PluginName
            Write-Host ("{0}: {1}" -f "Group Policy task ID $bgRunspaceID", "Retrieve GPO from $DomainName")`
            -ForegroundColor Magenta
            }
            Process{
                #Extract groups from domain
                $Domain = $ADObject.Domain
                $Connection = $ADObject.Domain.ADConnection
                $UseCredentials = $ADObject.UseCredentials
	            if($Connection){
		            $Searcher = New-Object System.DirectoryServices.DirectorySearcher -ArgumentList $Connection
			        $Searcher.SearchScope = "subtree"
                    $Searcher.SearchRoot = $Connection
                }
                if($Searcher){
                    $Searcher.PageSize = 200
                    #Retrieve all GPO which displayname like $DisplayName value
                    #$Searcher.Filter = "(&(objectClass=groupPolicyContainer)(displayname=$DisplayName))"
                    #Retrieve all GPO
                    $Searcher.Filter = '(objectClass=groupPolicyContainer)'
                    $GPos = $Searcher.FindAll()
                    foreach ($gpo in $GPos){
                        #https://msdn.microsoft.com/en-us/library/ms682264(v=vs.85).aspx
                        $TmpGPO = @{
			 				        "DisplayName" = $gpo.properties.displayname -join ''
							        "DistinguishedName" = $gpo.properties.distinguishedname -join ''
							        "CommonName" = $gpo.properties.cn -join ''
                                    "whenCreated" = $gpo.properties.whencreated -join ''
                                    "whenChanged" = $gpo.properties.whenchanged -join ''
							        "FilePath" = $gpo.properties.gpcfilesyspath -join ''
				        }
					    $AllGPO +=$TmpGPO 
                    }
                }  
            }
            End{
                if($AllGPO){
                    #Set PsObject from All Groups
                    $FormattedGPO = Set-PSObject -Object $AllGPO -type "AD.Arsenal.GroupPolicyObjects"
                    #Work with SyncHash
                    $SyncServer.$($PluginName)=$FormattedGPO

                    #Create custom object for store data
                    $GroupPolicyObjects = New-Object -TypeName PSCustomObject
                    $GroupPolicyObjects | Add-Member -type NoteProperty -name Data -value $FormattedGPO

                    #Formatting Excel Data
                    $Excelformatting = New-Object -TypeName PSCustomObject
                    $Excelformatting | Add-Member -type NoteProperty -name Data -value $FormattedGPO
                    $Excelformatting | Add-Member -type NoteProperty -name TableName -value "Group Policy"
                    $Excelformatting | Add-Member -type NoteProperty -name SheetName -value "Group Policy"
                    $Excelformatting | Add-Member -type NoteProperty -name isFreeze -value $True
                    $Excelformatting | Add-Member -type NoteProperty -name Type -value "CSVTable"

                    #Add Excel formatting into psobject
                    $GroupPolicyObjects | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting

                    #Add Group Policy data to object
                    $ReturnServerObject | Add-Member -type NoteProperty -name GroupPolicy -value $GroupPolicyObjects
            
                    #Add Groups data to object
                    $ReturnServerObject | Add-Member -type NoteProperty -name GroupPolicy -value $FormattedGPO

                    #Add Groups to report
                    $CustomReportFields = $ReturnServerObject.Report
                    $NewCustomReportFields = [array]$CustomReportFields+="GroupPolicy"
                    $ReturnServerObject | Add-Member -type NoteProperty -name Report -value $NewCustomReportFields -Force
                    #End
                }
            }