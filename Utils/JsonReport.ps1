Function Generate-Json{
[cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline = $True,ValueFromPipeLineByPropertyName = $True)]
        [Alias('IPAddress','Server','IPAddresses','Servers')]
        [Object]$ServerObject,
        
        [parameter()]
        [string]$RootPath

    )

    Begin{
        function ConvertTo-Json20([object] $AllItems){
            add-type -assembly system.web.extensions
            $JavaScriptSerializer = new-object system.web.script.serialization.javascriptSerializer
            $JavaScriptSerializer.MaxJsonLength = [System.Int32]::MaxValue
            $AllJsonResults = @()
            foreach ($item in $AllItems){
                $TmpDict = @{}
                $item.psobject.properties | Foreach { $TmpDict[$_.Name] = $_.Value }
                
                $AllJsonResults+=$JavaScriptSerializer.Serialize($TmpDict)
            }
            #Return Data
            return $AllJsonResults
        }
        Function Create-JsonFolderReport{
            [cmdletbinding()]
                Param (
                    [parameter(ValueFromPipeline = $True,ValueFromPipeLineByPropertyName = $True)]
                    [Alias('IPAddress','Server','IPAddresses','Servers')]
                    [String]$Computername,
        
                    [parameter()]
                    [string]$RootPath

                )
            $target = "$($RootPath)\JsonReport"
            if (!(Test-Path -Path $target)){
                $tmpdir = New-Item -ItemType Directory -Path $target
                Write-Host "Folder Reports created in $target...." -ForegroundColor Yellow
                return $target}
            else{
            Write-Host "Directory already exists...." -ForegroundColor Red
            return $target
            }
       }
       ##End of function
       if($ServerObject){
            Write-Host "Create report folder...." -ForegroundColor Green
            $ReportPath = Create-JsonFolderReport -RootPath $RootPath -Computername $ServerObject.ComputerName
            Write-Host "Report folder created in $($ReportPath)..." -ForegroundColor Green
       }
    }
    Process{
            if($ServerObject -and $ReportPath){
                Write-Host ("{0}: {1}" -f "JSON Task", "Generating XML report for data retrieved from $($Domain.Name)")`
                -ForegroundColor Magenta
                $ServerObject | %{
                    foreach ($query in $_.psobject.Properties){
                        if($query.Name -and $query.Value){
                            Write-Host "Export $($query.Name) to JSON file" -ForegroundColor Green
                            $JSONFile = ($ReportPath + "\" + ([System.Guid]::NewGuid()).ToString() +$query.Name+ ".json")
                            try{
                                if($query.value.Data){
                                    $output = ConvertTo-Json20 $query.value.Data
                                    Set-Content $JSONFile $output
                                    #$output | Out-File -FilePath $JSONFile
                                }
                            }
                            catch{
                                Write-Host "Function Generate-Json. Error in $($query.name)" -ForegroundColor Red 
                                Write-Host $_.Exception
                            }
                        }
                    }
                }
            }
        }
}