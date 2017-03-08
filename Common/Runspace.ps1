Function Get-RunSpaceADObject
{
    [CmdletBinding()]
    Param
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('IpAddress','Server')]
        [Object]
        $ComputerName=$env:computername,

        [Parameter(HelpMessage="Plugin to retrieve data",
                   Position=1)]
        [Array]
        $Plugins = $false,

        [Parameter(HelpMessage="Object with AD valuable data")]
        [Object]
        $ADObject = $false,
        
        [Parameter(HelpMessage="Maximum number of concurrent threads")]
        [ValidateRange(1,65535)]
        [int32]
        $Throttle = 32,
 
        [Parameter(HelpMessage="Timeout before a thread stops trying to gather the information")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,

        [Parameter(HelpMessage="Set this if you want run Single Query")]
        [switch]
        $SingleQuery,
        
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    Begin{

        #Function Get-RunSpaceData
        #http://learn-powershell.net/2012/05/10/speedy-network-information-query-using-powershell/
        Function Get-RunspaceData{
            [cmdletbinding()]
            param(
                [switch]$Wait,
                [String]$message = "Running Jobs"
            )
            Do {
                $more = $false
                $i = 0
                $total = $runspaces.Count        
                Foreach($runspace in $runspaces){
                    $StartTime = $runspacetimers[$runspace.ID]
                    $i++
                    If ($runspace.Runspace.isCompleted) {
                        #Write-Host $runspace.PluginName -ForegroundColor Yellow
                        $runspace.powershell.EndInvoke($runspace.Runspace)
                        $runspace.powershell.dispose()
                        $runspace.Runspace = $null
                        $runspace.powershell = $null                 
                    } ElseIf ($runspace.Runspace -ne $null) {
                        $more = $true
                    }
                }
                If ($more -AND $PSBoundParameters['Wait']) {
                    Start-Sleep -Milliseconds 100
                }   
                #Clean out unused runspace jobs
                $temphash = $runspaces.clone()
                $temphash | Where {
                    $_.runspace -eq $Null
                } | ForEach {
                    $Runspaces.remove($_)
                }  
                #Write-Progress -Activity $message -Status "Percent Complete" -PercentComplete $(($i/$total) * 100) 
            } while ($more -AND $PSBoundParameters['Wait'])
        }

        #Inicializar variables
        $SyncServer = [HashTable]::Synchronized(@{})
        $Global:ReturnServer = New-Object -TypeName PSCustomObject
        $runspacetimers = [HashTable]::Synchronized(@{})
        $runspaces = New-Object -TypeName System.Collections.ArrayList
        $Counter = 0
        $sessionstate = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($EntryVars in ('runspacetimers')){
            $sessionstate.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $EntryVars, (Get-Variable -Name $EntryVars -ValueOnly), ''))
        }
        $runspacepool = [runspacefactory]::CreateRunspacePool(1, $Throttle, $sessionstate, $Host)
        $runspacepool.ApartmentState = 'STA'
        $runspacepool.ThreadOptions = "ReuseThread"
        $runspacepool.Open()
    }
    Process{
        If($Plugins -ne $false){
            foreach ($Plugin in $Plugins){
                #Add plugin data to ADObject
                $PluginFullPath = $Plugin.FullName
                $PluginName = [io.path]::GetFileNameWithoutExtension($Plugin.FullName)
                Write-Verbose "Adding $PluginName to queue..."
                $NewPlugin = $ADObject | Select-Object *
                $NewPlugin | Add-Member -type NoteProperty -name PluginName -value $PluginName -Force
                #End plugin work
                $ScriptBlockPlugin = [ScriptBlock]::Create($(Get-Content $PluginFullPath | Out-String))
                $Counter++
                $PowerShell = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlockPlugin)
                $null = $PowerShell.AddParameter('bgRunspaceID',$Counter)
                $null = $PowerShell.AddParameter('SyncServer',$SyncServer)
                $null = $PowerShell.AddParameter('ADObject',$NewPlugin)
                $null = $PowerShell.AddParameter('ReturnServerObject',$ReturnServer)
                $PowerShell.RunspacePool = $runspacepool

                [void]$runspaces.Add(@{
                    runspace = $PowerShell.BeginInvoke()
                    PowerShell = $PowerShell
                    Computer = $ComputerName.ComputerName
                    PluginName = $PluginName
                    ID = $Counter
                    })
            }
        }
        Get-RunspaceData 
    }
    End{
        Get-RunspaceData -Wait
        $runspacepool.Close()
        $runspacepool.Dispose()
        #$SyncServer
        return $ReturnServer
    }
}


