Function Run-Jobs{
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

        [Parameter(HelpMessage="Plugin to retrieve data")]
        [Array]
        $Plugins = $false,

        [Parameter(HelpMessage="Object with AD data")]
        [Object]
        $ADObject = $false,
        
        [Parameter(HelpMessage="Maximum number of max jobs")]
        [ValidateRange(1,65535)]
        [int32]
        $MaxJobs = 10,

        [Parameter(HelpMessage="Set this if you want run Single Query")]
        [switch]
        $SingleQuery
    )
    Begin{
        Function Global:Get-DataFromAD{
            $VerbosePreference = 'continue'
            If ($queue.count -gt 0) {
                Write-Verbose ("Queue Count: {0}" -f $queue.count)
                $TaskObject = $queue.Dequeue() 
                #$TaskObject | fl
                $ScriptBlockPlugin = [ScriptBlock]::Create($(Get-Content $TaskObject.Plugin | Out-String))
                $TaskName = $TaskObject.PluginName
                $Plugin = $TaskObject.Plugin
                #Create task
                $NewTask = Start-Job -Name $TaskName -FilePath $Plugin -ArgumentList $TaskObject

                #Register Event
                Register-ObjectEvent -InputObject $NewTask -EventName StateChanged -Action {
                    #Set verbose to continue to see the output on the screen
                    $VerbosePreference = 'continue'
                    $TaskUpdate = $eventsubscriber.sourceobject.name      
                    $Global:Data += Receive-Job -Job $eventsubscriber.sourceobject
                    Write-Verbose "Removing: $($eventsubscriber.sourceobject.Name)"           
                    Remove-Job -Job $eventsubscriber.sourceobject
                    Write-Verbose "Unregistering: $($eventsubscriber.SourceIdentifier)"
                    Unregister-Event $eventsubscriber.SourceIdentifier
                    Write-Verbose "Removing: $($eventsubscriber.SourceIdentifier)"
                    Remove-Job -Name $eventsubscriber.SourceIdentifier
                    Remove-Variable results
                    If ($queue.count -gt 0 -OR (Get-Job)){
                        Write-Verbose "Running Rjob"
                        Get-DataFromAD
                    }ElseIf (-NOT (Get-Job)) {
                        $End = (Get-Date)        
                        $timeToCleanup = ($End - $Start).TotalMilliseconds    
                        Write-Host ('{0,-30} : {1,10:#,##0.00} ms' -f 'Time to exit background job', $timeToCleanup) -fore Green -Back Black
                        Write-Host -ForegroundColor Green "Check the `$Data variable for report of online/offline systems"
                    }           
                } | Out-Null
                Write-Verbose "Created Event for $($NewTask.Name)"
            }
        }        
    }
    Process{
        #Capture date
        $Start = Get-Date
        if($Plugins -ne $false -and $ADObject){
            #Queue the items up
            $Global:queue = [System.Collections.Queue]::Synchronized( (New-Object System.Collections.Queue) )
            foreach ($plugin in $Plugins){
                $PluginName = [io.path]::GetFileNameWithoutExtension($plugin)
                Write-Verbose "Adding $PluginName to queue..."
                $NewPlugin = $ADObject | Select-Object *
                $NewPlugin | Add-Member -type NoteProperty -name PluginName -value $PluginName -Force
                $NewPlugin | Add-Member -type NoteProperty -name Plugin -value $plugin -Force
                #Add PsObject to Queue
                $queue.Enqueue($NewPlugin)
                #$NewPlugin | fl
            }
            If ($queue.count -lt $MaxJobs) {
                $MaxJobs = $queue.count
            }
            # Start up to the max number of concurrent jobs
            # Each job will take care of running the rest
            for( $i = 0; $i -lt $MaxJobs; $i++ ) {
                Get-DataFromAD
            } 
        }
    }
    End{
        #Nothing to do here
    }



}