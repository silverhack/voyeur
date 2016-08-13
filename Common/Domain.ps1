#---------------------------------------------------
# Return the Domain SID
#---------------------------------------------------

function Get-DomainSID{
    try{
        $KerberosEncType = [System.DirectoryServices.AuthenticationTypes]::Sealing -bor [System.DirectoryServices.AuthenticationTypes]::Secure 
        $SSLEncType = [System.DirectoryServices.AuthenticationTypes]::SecureSocketsLayer
        $GlobalCatalog = $Global:Forest.FindGlobalCatalog()
		if ($Global:AllDomainData -and $GlobalCatalog){
            $newLDAPConnection = $Global:AllDomainData.LDAPBase
            if($UseSSL){
                $DN = "GC://$($GlobalCatalog.Name):636/$($Global:AllDomainData.DistinguishedName)"
                $newLDAPConnection = ,$DN+$newLDAPConnection
                $ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $newLDAPConnection
            }
            elseif(-NOT $UseSSL){
                $DN = "GC://$($GlobalCatalog.IPAddress)/$($Global:AllDomainData.DistinguishedName)"
                $newLDAPConnection = ,$DN+$newLDAPConnection
                $ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $newLDAPConnection
                $ADObject.AuthenticationType = $KerberosEncType
            }
		}
        if($ADObject){
            #$ADObject | gm
            $DomainSid = New-Object System.Security.Principal.SecurityIdentifier($ADObject.objectSid[0], 0)
            return $DomainSid.ToString()
        }
     }
    catch
        {
            Write-Host "Function Get-DomainSID $($_.Exception.Message)" -ForegroundColor Red
        }     
        
}

#---------------------------------------------------
# Return the Current Forest
#---------------------------------------------------

function Get-CurrentForest{
    try{
        if($Global:AllDomainData){
            $MyContext = $AllDomainData.DomainContext
            #Forest Level Connection
            if($UseCredentials){
                $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest", $MyContext.forest,$($credential.UserName),$($credential.GetNetworkCredential().password))
            }
            else{
                $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest", $MyContext.forest)
            }
		    
		    $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)

            return $forest
        }
    }
    Catch{
        Write-Host "Function Get-CurrentForest $($_.Exception.Message)" -ForegroundColor Red
    }
}

#---------------------------------------------------
# Try to return domain passed variable
#---------------------------------------------------
function Get-DomainInfo{
    <#
    try { 
        $assemType = 'System.DirectoryServices.AccountManagement'
        $assem = [reflection.assembly]::LoadWithPartialName($assemType)
    }
    catch{
        throw "Failed to load assembly System.DirectoryServices.AccountManagement"
    }
    #>
    try{
        #Initialize array and create new object
        $isConnected = $false
        $LDAPRootDSE = $null
        $LDAP = $null
        $newLDAPConnection = @()
        $DomainObject = New-Object -TypeName PSCustomObject
        $KerberosEncType = [System.DirectoryServices.AuthenticationTypes]::Sealing -bor [System.DirectoryServices.AuthenticationTypes]::Secure 
        $SSLEncType = [System.DirectoryServices.AuthenticationTypes]::SecureSocketsLayer
        if($Global:DomainName -and (-NOT $Global:UseSSL)){
            $LDAP = "LDAP://{0}" -f $DomainName
            $LDAPRootDSE = "LDAP://{0}/RootDSE" -f $DomainName
        }
        elseif($Global:DomainName -and $Global:UseSSL){
            $LDAPRootDSE = "LDAP://{0}:636/RootDSE" -f $DomainName
            $LDAP = "LDAP://{0}:636" -f $DomainName
            $UseSSL = $true
        }
        elseif(!$Global:DomainName -and $Global:UseSSL){
            $CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $LDAPRootDSE = "LDAP://{0}:636/RootDSE" -f $CurrentDomain.name
            $LDAP = "LDAP://{0}:636" -f $CurrentDomain.name
            $UseSSL = $true
        }
        else{
            $CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $LDAPRootDSE = "LDAP://{0}/RootDSE" -f $CurrentDomain.name
            $LDAP = "LDAP://{0}" -f $CurrentDomain.name
            
        }
        #Add LDAP connection and auth to array
        if($UseCredentials){
            $newLDAPConnection+=$credential.UserName
            $newLDAPConnection+=$credential.GetNetworkCredential().password
        }        
        $Conn = $newLDAPConnection.clone()
        $Conn = ,$LDAP+$Conn
        
        #Prepare new connection
        $DomainConnection = New-Object -TypeName System.DirectoryServices.DirectoryEntry($Conn)
        if(-NOT $UseSSL){$DomainConnection.AuthenticationType = $KerberosEncType}
        #elseif($UseSSL){$DomainConnection.AuthenticationType = $SSLEncType}
        
        
        #Prepare new RootDSE Connection
        $RootConn = $newLDAPConnection.clone()
        $RootConn = ,$LDAPRootDSE+$RootConn
        $RootDSEConnection = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $RootConn
        if(-NOT $UseSSL){$RootDSEConnection.AuthenticationType = $KerberosEncType}
        #elseif($UseSSL){$RootDSEConnection.AuthenticationType = $SSLEncType}
        
        $DomainObject | Add-Member -type NoteProperty -name LDAPBase -value $newLDAPConnection
        $DomainObject | Add-Member -type NoteProperty -name LDAPConnection -value $Conn
        $DomainObject | Add-Member -type NoteProperty -name RootDSELDAPConnection -value $RootConn
        
        #Check if connected
        if($DomainConnection.Name -and $RootDSEConnection.Name){
            $isConnected = $true
            write-host "Successfully connect to $($DomainConnection.distinguishedname)" -ForegroundColor Green
            
            #Extract NetBios Name
            # Connect to the Configuration Naming Context
            $LDAPConfigurationNamingContext = "{0}/{1}" -f $($LDAP), $RootDSEConnection.Get("configurationNamingContext")
            $NewLDAP = $newLDAPConnection.clone()
            $NewLDAP = ,$LDAPConfigurationNamingContext+$NewLDAP
            
            $configSearchRoot = New-Object -TypeName System.DirectoryServices.DirectoryEntry ($NewLDAP)
            if(-NOT $UseSSL){$configSearchRoot.AuthenticationType = $KerberosEncType}
            #elseif($UseSSL){$configSearchRoot.AuthenticationType = $SSLEncType}
                        
            Write-Host "Connected to $($configSearchRoot.distinguishedName)" -ForegroundColor Green
            
            # Configure the filter
            $filter = "(NETBIOSName=*)"
            # Search for all partitions where the NetBIOSName is set
            $configSearch = New-Object DirectoryServices.DirectorySearcher($configSearchRoot, $filter)
            # Configure search to return dnsroot and ncname attributes
            $retVal = $configSearch.PropertiesToLoad.Add("dnsroot")
            $retVal = $configSearch.PropertiesToLoad.Add("ncname")
            $retVal = $configSearch.PropertiesToLoad.Add("name")

            #Search data
            $Data = $configSearch.FindAll()
            if($Data){
                foreach ($info in $Data){
                    $DomainObject | Add-Member -type NoteProperty -name Name -value $info.Properties.Item("name")[0]
                    $DomainObject | Add-Member -type NoteProperty -name DistinguishedName -value $info.Properties.Item("ncname")[0]
                    $DomainObject | Add-Member -type NoteProperty -name DNSRoot -value $info.Properties.Item("dnsroot")[0]
                }
            }
            
            #Add values to object
            $DomainObject | Add-Member -type NoteProperty -name DomainConnection -value $DomainConnection
            $DomainObject | Add-Member -type NoteProperty -name RootDSEConnection -value $RootDSEConnection
            $DomainObject | Add-Member -type NoteProperty -name DNSHostName -value $RootDSEConnection.dnshostname.toString()
            
            #Get Domain context
            $NewConn = $newLDAPConnection.clone()
            $NewConn = ,$DomainObject.Name+$NewConn
            $NewConn = ,"Domain"+$NewConn
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(`
                        $NewConn)
            $CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext) 
            
            #Add values to object
            $DomainObject | Add-Member -type NoteProperty -name DomainContext -value $CurrentDomain
            $DomainObject | Add-Member -type NoteProperty -name DomainFunctionalLevel -value $CurrentDomain.DomainMode
            $DomainObject | Add-Member -type NoteProperty -name LDAP -value $LDAP
            $DomainObject | Add-Member -type NoteProperty -name LDAPRootDSE -value $LDAPRootDSE

            #Prepare Domain Connection
            $MyConn = $newLDAPConnection.clone()
            if($Global:SearchRoot){
                $DN = "{0}/{1}" -f $LDAP,$Global:SearchRoot
                $MyConn = ,$DN+$MyConn
            }
            else{
                $DN = "{0}" -f $LDAP
                $MyConn = ,$DN+$MyConn
            }
            $MyADConnection = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $MyConn
            if(-NOT $UseSSL){
                $MyADConnection.AuthenticationType = $KerberosEncType
            }
            $DomainObject | Add-Member -type NoteProperty -name ADConnection -value $MyADConnection
            
            #>
            $DomainConnection.Close()
            $RootDSEConnection.Close()
            $configSearchRoot.Close()

        }
        return $DomainObject
        
    }
    catch{
        Write-Host "Function Get-DomainInfo $($_.Exception)" -ForegroundColor Red
        break
    }
}
