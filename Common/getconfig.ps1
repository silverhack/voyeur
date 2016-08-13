Function Get-VoyeurConf{
    Param (
        [parameter(Mandatory=$true, HelpMessage="Path to voyeur config file")]
        [string]$Path,
        [parameter(Mandatory=$true, HelpMessage="Path to voyeur config file")]
        [string]$Node
    )
    Begin{
        $global:appSettings = @{}
        try{
            [xml]$config = Get-Content $Path.ToString()
            foreach ($addNode in $config.configuration.$($Node).add){
                if ($addNode.Value.Contains(‘,’)){
                    # Array case
                    $value = $addNode.Value.Split(‘,’)
                    for ($i = 0; $i -lt $value.length; $i++){ 
                        $value[$i] = $value[$i].Trim() 
                    }
                }
                else{
                    # Scalar case
                    $value = $addNode.Value
                }
                $global:appSettings[$addNode.Key] = $value
            }
        }
        catch{
            throw ("{0}: {1}" -f "Fatal error in config file",$_.Exception.Message)
        }
    }
    Process{
        if($global:appSettings.Count -ge 0){            return $global:appSettings        }
        else{
            return $false
        }
    }
    End{
        #Nothing to do here
    }
}