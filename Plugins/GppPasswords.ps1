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
            #define helper function that decodes and decrypts password
            function Get-DecryptedCpassword {
                Param (
                    [string] $Cpassword 
                )

                try {
                    #Append appropriate padding based on string length  
                    $Mod = ($Cpassword.length % 4)
                    if ($Mod -ne 0) {$Cpassword += ('=' * (4 - $Mod))}

                    $Base64Decoded = [Convert]::FromBase64String($Cpassword)
            
                    #Create a new AES .NET Crypto Object
                    $AesObject = New-Object System.Security.Cryptography.AesCryptoServiceProvider
                    [Byte[]] $AesKey = @(0x4e,0x99,0x06,0xe8,0xfc,0xb6,0x6c,0xc9,0xfa,0xf4,0x93,0x10,0x62,0x0f,0xfe,0xe8,
                                         0xf4,0x96,0xe8,0x06,0xcc,0x05,0x79,0x90,0x20,0x9b,0x09,0xa4,0x33,0xb6,0x6c,0x1b)
            
                    #Set IV to all nulls to prevent dynamic generation of IV value
                    $AesIV = New-Object Byte[]($AesObject.IV.Length) 
                    $AesObject.IV = $AesIV
                    $AesObject.Key = $AesKey
                    $DecryptorObject = $AesObject.CreateDecryptor() 
                    [Byte[]] $OutBlock = $DecryptorObject.TransformFinalBlock($Base64Decoded, 0, $Base64Decoded.length)
            
                    return [System.Text.UnicodeEncoding]::Unicode.GetString($OutBlock)
                } 
        
                catch {Write-Error $Error[0]}
            }  
            #Check if part of domain
            function Check-IfDomain{             
                [CmdletBinding()]             
                param (             
                [parameter(Position=0,            
                   Mandatory=$true,            
                   ValueFromPipeline=$true,             
                   ValueFromPipelineByPropertyName=$true)]            
                   [string]$computer="."             
                )                         
                PROCESS{            
                 $DomainData = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computer |            
                 select Name, Domain, partofdomain

                 #Return Object
                 return $DomainData                    
                }#process
            }
            #Get potential files containing passwords from Group Policy objects
            Function Get-XMLFiles{
                $AllFiles = $false
                if($ADObject.Domain.DistinguishedName -or $ADObject.Domain.DNSRoot -or $ADObject.Domain.DNSHostName){
                    try{
                        $DomainName = $ADObject.Domain.DNSRoot
                        $AllFiles = Get-ChildItem "\\$DomainName\SYSVOL" -Recurse -ErrorAction Stop -Include 'Groups.xml','Services.xml','Scheduledtasks.xml','DataSources.xml','Drives.xml'
                        #Push-Location
                        #Start-Sleep -Milliseconds 100
                        #Pop-Location
                    }
                    catch [System.Management.Automation.ItemNotFoundException]{
                        Write-Warning "Item not found"
                    }
                    catch [System.UnauthorizedAccessException]{
                        Write-Warning "Access denied"
                    }
                    catch{
                        Write-Host "Function Get-XMLFiles $($_.Exception.Message)" -ForegroundColor Red 
                    }
                }
                return $AllFiles
            }
        }
        Process{
            #Start GPP plugin
            $DomainName = $ADObject.Domain.Name
            $PluginName = $ADObject.PluginName
            Write-Host ("{0}: {1}" -f "Group Policy task ID $bgRunspaceID", "Retrieve potential files with Username:passwords from $DomainName")`
            -ForegroundColor Magenta
            $AllItems = Get-XMLFiles
            if($AllItems){
                $GPPLeaks = @()
                foreach ($File in $AllItems){
                    $Filename = $File.Name
                    $Filepath = $File.VersionInfo.FileName
                    #put filename in $XmlFile
                    [xml] $Xml = Get-Content($File)
			
                    #declare blank variables
                    $Cpassword = ''
                    $UserName = ''
                    $NewName = ''
                    $Changed = ''
                 
                    switch ($Filename) {

                        'Groups.xml' {
                            $Cpassword = $Xml.Groups.User.Properties.cpassword
                            $UserName = $Xml.Groups.User.Properties.userName
                            $NewName = $Xml.Groups.User.Properties.newName
                            $Changed = $Xml.Groups.User.changed
                        }
        
                        'Services.xml' {
                            $Cpassword = $Xml.NTServices.NTService.Properties.cpassword
                            $UserName = $Xml.NTServices.NTService.Properties.accountName
                            $Changed = $Xml.NTServices.NTService.changed
                        }
        
                        'Scheduledtasks.xml' {
                            $Cpassword = $Xml.ScheduledTasks.Task.Properties.cpassword
                            $UserName = $Xml.ScheduledTasks.Task.Properties.runAs
                            $Changed = $Xml.ScheduledTasks.Task.changed
                        }
        
                        'DataSources.xml' {
                            $Cpassword = $Xml.DataSources.DataSource.Properties.cpassword
                            $UserName = $Xml.DataSources.DataSource.Properties.username
                            $Changed = $Xml.DataSources.DataSource.changed
                        }
				
				        'Drives.xml' {
                            $Cpassword = $Xml.Drives.Drive.Properties.cpassword
                            $UserName = $Xml.Drives.Drive.Properties.username
                            $Changed = $Xml.Drives.Drive.changed
                        }
                    }
                    if ($Cpassword) {$Password = Get-DecryptedCpassword $Cpassword}
                    else {Write-Verbose "No encrypted passwords found in $Filepath"}
                    #Create custom object to output results
                    $ObjectProperties = @{'Password' = [String]$Password;
                                          'UserName' = [String]$UserName;
                                          'Changed' = [String]$Changed;
                                          'NewName' = [String]$NewName
                                          'File' = [String]$Filepath}
                
                    $ResultsObject = New-Object -TypeName PSObject -Property $ObjectProperties
			        $ResultsObject.PSObject.typenames.insert(0,'Arsenal.AD.GetPasswords')
                    $GPPLeaks += $ResultsObject
                }
            }
            else{
                Write-Warning 'No group policy preferences found on Domain....'
                Break
            }
        }
        End{
            #Work with SyncHash
            $SyncServer.$($PluginName)=$GPPLeaks
            #Add Group policy passwords data to object

            #Create custom object for store data
            $GroupPolicyLeaks = New-Object -TypeName PSCustomObject
            $GroupPolicyLeaks | Add-Member -type NoteProperty -name Data -value $GPPLeaks

            #Formatting Excel Data
            $Excelformatting = New-Object -TypeName PSCustomObject
            $Excelformatting | Add-Member -type NoteProperty -name Data -value $GPPLeaks
            $Excelformatting | Add-Member -type NoteProperty -name TableName -value "GPPLeaks"
            $Excelformatting | Add-Member -type NoteProperty -name SheetName -value "GPPLeaks"
            $Excelformatting | Add-Member -type NoteProperty -name isFreeze -value $True
            $Excelformatting | Add-Member -type NoteProperty -name Type -value "CSVTable"

            #Add Excel formatting into psobject
            $GroupPolicyLeaks | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting

            #Add Groups data to object
            if($GPPLeaks){
                $ReturnServerObject | Add-Member -type NoteProperty -name GppLeaks -value $GroupPolicyLeaks
            }

            #Add GPP to report
            $CustomReportFields = $ReturnServerObject.Report
            $NewCustomReportFields = [array]$CustomReportFields+="GPPleaks"
            $ReturnServerObject | Add-Member -type NoteProperty -name Report -value $NewCustomReportFields -Force
            #End
        }