Function Generate-CSV{
[cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline = $True,ValueFromPipeLineByPropertyName = $True)]
        [Object]$ServerObject,
        
        [parameter()]
        [string]$RootPath

    )

    Begin{
          Function Create-CSVFolderReport{
            [cmdletbinding()]
                Param (
                    [parameter(ValueFromPipeline = $True,ValueFromPipeLineByPropertyName = $True)]
                    [Alias('IPAddress','Server','IPAddresses','Servers')]
                    [String]$Computername,
        
                    [parameter()]
                    [string]$RootPath

                )
            $target = "$($RootPath)\CSVReport"
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
    }
    Process{
            if($ServerObject){
                Write-Host "Create report folder...." -ForegroundColor Green
                $ReportPath = Create-CSVFolderReport -RootPath $RootPath -Computername $ServerObject.ComputerName
                Write-Host "Report folder created in $($ReportPath)..." -ForegroundColor Green
            }
            if($ServerObject -and $ReportPath){
                Write-Host ("{0}: {1}" -f "CSV Task", "Generating CSV report for data retrieved from $($Domain.Name)")`
                -ForegroundColor Magenta
                $ServerObject | %{
                    foreach ($query in $_.psobject.Properties){
                        if($query.Name -and $query.Value){
                            Write-Host "Export $($query.Name) to CSV file" -ForegroundColor Green
                            $CSVFile = ($ReportPath + "\" + ([System.Guid]::NewGuid()).ToString() +$query.Name+ ".csv")
                            try{
                                if($query.value.Data){
                                    $query.value.Data | Export-Csv -NoTypeInformation -Path $CSVFile 
                                }
                            }
                            catch{
                                Write-Host "Function Generate-CSV. Error in $($query.name)" -ForegroundColor Red
                            }
                        }
                    }
                }
           }             
    }
    End{
        #Nothing to do here
    }
} 