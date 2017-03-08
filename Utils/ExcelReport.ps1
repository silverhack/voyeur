Function Generate-Excel{
[cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline = $True,ValueFromPipeLineByPropertyName = $True)]
        [Alias('ServerObject')]
        [Object]$AllData,
        
        [parameter(ValueFromPipeline = $True,ValueFromPipeLineByPropertyName = $True)]
        [Alias('Formatting')]
        [Object]$TableFormatting,

        [parameter(ValueFromPipeline = $True,ValueFromPipeLineByPropertyName = $True)]
        [Alias('Style')]
        [Object]$HeaderStyle,

        [parameter(ValueFromPipeline = $True,ValueFromPipeLineByPropertyName = $True)]
        [Alias('Settings')]
        [Object]$ExcelSettings,

        [parameter()]
        [string]$RootPath

    )
    Begin{
        Function Create-ExcelFolderReport{
                [cmdletbinding()]
                Param (
                    [parameter()]
                    [string]$RootPath
                )
            $target = "$($RootPath)\ExcelReport"
            if (!(Test-Path -Path $target)){
                $tmpdir = New-Item -ItemType Directory -Path $target
                Write-Host "Folder Reports created in $target...." -ForegroundColor Yellow
                return $target}
            else{
            Write-Host "Directory already exists...." -ForegroundColor Red
            return $target
            }
       }   
            
    }
    Process{
        if($AllData -and $ExcelSettings -and $TableFormatting){
            #Add Types
			Add-Type -AssemblyName Microsoft.Office.Interop.Excel
            #Create Excel object
            $isDebug = [System.Convert]::ToBoolean($ExcelSettings["Debug"])
            Create-Excel -ExcelDebugging $isDebug
            if($ExcelSettings){
                # Create About Page
			    Create-About -ExcelSettings $ExcelSettings
            }
            #Get Language Settings
            $Language = Get-ExcelLanguage 
            if ($Language){
                #Try to get config data
                if($HeaderStyle["$Language"]){
                    #Detected language
                    Write-Host "Locale detected..." -ForegroundColor Magenta
                    $FormatTable = $TableFormatting["Style"]
                    $Header = $HeaderStyle["$Language"]
                }
                else{
                    $FormatTable = $false
                    $Header = $false
                }
            }
            #Populate data into Excel sheets
            $AllData | % {
                Foreach ($newDataSheet in $_.psobject.Properties){
                    if($newDataSheet.Value.Excelformat){
                        Write-Host "Add $($newDataSheet.Name) data into Excel..." -ForegroundColor Magenta
                        #Extract format of sheet (New Table, CSV2Table, Chart, etc...)
                        $TypeOfData = $newDataSheet.Value.ExcelFormat.Type
                        #Switch cases
                        switch($TypeOfData){
                            'CSVTable'
                            {
                              $Data =  $newDataSheet.Value.ExcelFormat.Data
                              $Title = $newDataSheet.Value.ExcelFormat.SheetName
                              $TableTitle = $newDataSheet.Value.ExcelFormat.TableName
                              $freeze = $newDataSheet.Value.ExcelFormat.isFreeze
                              if($newDataSheet.Value.ExcelFormat.IconSet){
                                $columnName = $newDataSheet.Value.ExcelFormat.IconColumnName
                              }
                              else{$columnName=$null}
                              Create-CSV2Table -Data $Data -Title $Title -TableTitle $TableTitle `
                                               -TableStyle $FormatTable -isFreeze $freeze `
                                               -iconColumnName $columnName| Out-Null
                            }
                            'NewTable'
                            { 
                              #Extract data from psobject  
                              $Data =  $newDataSheet.Value.ExcelFormat.Data
                              $ShowHeaders = $newDataSheet.Value.ExcelFormat.showHeaders
                              $MyHeaders = $newDataSheet.Value.ExcelFormat.addHeader
                              $ShowTotals = $newDataSheet.Value.ExcelFormat.showTotals
                              $Position = $newDataSheet.Value.ExcelFormat.position
                              $Title = $newDataSheet.Value.ExcelFormat.SheetName
                              $TableName = $newDataSheet.Value.ExcelFormat.TableName
                              $isnewSheet = [System.Convert]::ToBoolean($newDataSheet.Value.ExcelFormat.isnewSheet)
                              $addNewChart = [System.Convert]::ToBoolean($newDataSheet.Value.ExcelFormat.addChart)
                              $ChartType = $newDataSheet.Value.ExcelFormat.chartType
                              $HasDatatable = [System.Convert]::ToBoolean($newDataSheet.Value.ExcelFormat.hasDataTable)
                              $chartStyle = $newDataSheet.Value.ExcelFormat.Style
                              $chartTitle = $newDataSheet.Value.ExcelFormat.ChartTitle
     
                              #Create new table with data
                              Create-Table -ShowTotals $ShowTotals -ShowHeaders $ShowHeaders -Data $Data `
                                           -SheetName $Title -TableTitle $TableName -Position $Position `
                                           -Header $MyHeaders -isNewSheet $isnewSheet -addNewChart $addNewChart `
                                           -ChartType $ChartType -ChartTitle $chartTitle `
                                           -ChartStyle $chartStyle -HasDataTable $HasDatatable `
                                           -HeaderStyle $Header | Out-Null
                            }
                        }
                    }
                }
            }
        }
        #Delete Sheet1 and create index
		$Excel.WorkSheets.Item($Excel.WorkSheets.Count).Delete() | Out-Null
        Create-Index -ExcelSettings $ExcelSettings
    }
    End{
        #Create Report Folder
        Write-Host "Creating report folder...." -ForegroundColor Green
        $ReportPath = Create-ExcelFolderReport -RootPath $RootPath
        Write-Host "Report folder created in $($ReportPath)..." -ForegroundColor Green

        #Save Excel
        Save-Excel -Path $ReportPath
        #Release Excel Object
        Release-ExcelObject 
    }
}