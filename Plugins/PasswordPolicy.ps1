#Plugin extract Domain password policy from AD
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
            #Start computers plugin
            $DomainName = $ADObject.Domain.Name
            $PluginName = $ADObject.PluginName
            Write-Host ("{0}: {1}" -f "Domain Password Policy task ID $bgRunspaceID", "Retrieve info data data from $DomainName")`
            -ForegroundColor Magenta
            }
        Process{
            #Extract Domain password policy from domain
            $Domain = $ADObject.Domain.DomainContext.GetDirectoryEntry()
            $UseCredentials = $ADObject.UseCredentials
	        if($Domain){
                $DomainDetailmaxPwdAgeValue = $Domain.maxPwdAge.Value 
                $DomainDetailminPwdAgeValue = $Domain.minPwdAge.Value 
                $DomainDetailmaxPwdAgeInt64 = $Domain.ConvertLargeIntegerToInt64($DomainDetailmaxPwdAgeValue) 
                $DomainDetailminPwdAgeInt64 = $Domain.ConvertLargeIntegerToInt64($DomainDetailminPwdAgeValue) 

                $MaxPwdAge = -$DomainDetailmaxPwdAgeInt64/(600000000 * 1440) 
                $MinPwdAge = -$DomainDetailminPwdAgeInt64/(600000000 * 1440)

                $DomainDetailminPwdLength = $Domain.minPwdLength 
                $DomainDetailpwdHistoryLength = $Domain.pwdHistoryLength
                #Construct object
                $DomainPasswordPolicy = @{
			 				"MaxPasswordAge" = $MaxPwdAge
							"MinPasswordAge" = $MinPwdAge
							"MinPasswordLength" = $DomainDetailminPwdLength.ToString()
							"PasswordHistoryLength" = $DomainDetailpwdHistoryLength.ToString()
				}   
            }
                
        }
        End{
            #Set PsObject from Password Policy object
            $DPP = Set-PSObject -Object $DomainPasswordPolicy -type "AD.Arsenal.DomainPasswordPolicy"
            #Work with SyncHash
            $SyncServer.$($PluginName)=$DPP

            #Create custom object for store data
            $DomainPasswordPolicy = New-Object -TypeName PSCustomObject
            $DomainPasswordPolicy | Add-Member -type NoteProperty -name Data -value $DPP

            #Formatting Excel Data
            $Excelformatting = New-Object -TypeName PSCustomObject
            $Excelformatting | Add-Member -type NoteProperty -name Data -value $DPP
            $Excelformatting | Add-Member -type NoteProperty -name TableName -value "Domain Password Policy"
            $Excelformatting | Add-Member -type NoteProperty -name SheetName -value "Domain Password Policy"
            $Excelformatting | Add-Member -type NoteProperty -name isFreeze -value $True
            $Excelformatting | Add-Member -type NoteProperty -name Type -value "CSVTable"

            #Add Excel formatting into psobject
            $DomainPasswordPolicy | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting
            
            #Add DPP data to object
            if($DPP){
                $ReturnServerObject | Add-Member -type NoteProperty -name DomainPasswordPolicy -value $DomainPasswordPolicy
            }

            #Add DPP to report
            $CustomReportFields = $ReturnServerObject.Report
            $NewCustomReportFields = [array]$CustomReportFields+="DomainPasswordPolicy"
            $ReturnServerObject | Add-Member -type NoteProperty -name Report -value $NewCustomReportFields -Force
            #End
        }
            