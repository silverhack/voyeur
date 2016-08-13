Function Generate-XML{
[cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline = $True,ValueFromPipeLineByPropertyName = $True)]
        [Alias('IPAddress','Server','IPAddresses','Servers')]
        [Object]$ServerObject,
        
        [parameter()]
        [string]$RootPath

    )

    Begin{
        Function Create-XMLFolderReport{
            [cmdletbinding()]
                Param (
                    [parameter(ValueFromPipeline = $True,ValueFromPipeLineByPropertyName = $True)]
                    [Alias('IPAddress','Server','IPAddresses','Servers')]
                    [String]$Computername,
        
                    [parameter()]
                    [string]$RootPath

                )
            $target = "$($RootPath)\XMLReport"
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
            $ReportPath = Create-XMLFolderReport -RootPath $RootPath -Computername $ServerObject.ComputerName
            Write-Host "Report folder created in $($ReportPath)..." -ForegroundColor Green
       }
    }
    Process{
            if($ServerObject -and $ReportPath){
                Write-Host ("{0}: {1}" -f "XML Task", "Generating XML report for data retrieved from $($Domain.Name)")`
                -ForegroundColor Magenta
                $ServerObject | %{
                    foreach ($query in $_.psobject.Properties){
                        if($query.Name -and $query.Value){
                            Write-Host "Export $($query.Name) to XML file" -ForegroundColor Green
                            $XMLFile = ($ReportPath + "\" + ([System.Guid]::NewGuid()).ToString() +$query.Name+ ".xml")
                            try{
                                if($query.value.Data){
                                    ($query.value.Data | ConvertTo-Xml).Save($XMLFile)
                                }
                            }
                            catch{
                                Write-Host "Function Generate-XML. Error in $($query.name)" -ForegroundColor Red
                            }
                        }
                    }
                }
            }
        }
}