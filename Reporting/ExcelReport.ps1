<#
.SYNOPSIS
	The script scan Active Directory and retrieve huge data of many sources like:
	-User objects
	-Computer objects
	-Organizational Unit
	-Forest
	-Access control list
	-Etc.
	
	Ability to export data to CSV and Excel format. Default save format is CSV
	
.NOTES
	Author		: Juan Garrido (http://windowstips.wordpress.com)
	Twitter		: @tr1ana
	Company		: http://www.innotecsystem.com
	File Name	: voyeur.ps1	

#>

#---------------------------------------------------
# Function to create OBJ Excel
#---------------------------------------------------

function global:Create-Excel ()
	{
		#Create Excel 
		[Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'
		$Excel = new-object -com Excel.Application
		$Excel.visible = $true
		$objWorkBook = $Excel.WorkBooks.Add()
	#$objWorkBook.Sheets | Where-Object {$_ -ne "Sheet1"}
		#Remove WorkSheets
		1..2 | ForEach {
	    $objWorkbook.WorkSheets.Item(2).Delete()
		}
		return $objWorkBook,$Excel
	}
	
#---------------------------------------------------
# Function to create WorkSheet
#---------------------------------------------------

function Create-WorkSheet ([Object] $WorkBook, [String] $Title, [Object] $Excel)
	{
		$WorkSheet = $WorkBook.Worksheets.Add()
		$WorkSheet.Activate() | Out-Null
		$WorkSheet.Name = $Title
		$WorkSheet.Select()
		$Excel.ActiveWindow.Displaygridlines = $false 
		return $WorkSheet
	}

#---------------------------------------------------
# Function to create table through CSV data
#---------------------------------------------------

function Create-CSV2Table ([Object] $Data, [String] $Title, [String] $TableTitle, [Object] $WorkBook, [Object] $Excel, [Bool] $isFreeze)
	{
		if ($Data -ne $null)
			{
				$CSVFile = ($env:temp + "\" + ([System.Guid]::NewGuid()).ToString() + ".csv")
				$Data | Export-Csv -path $CSVFile -noTypeInformation
				$WorkSheet = Create-WorkSheet $WorkBook $Title $Excel
				#Define the connection string and where the data is supposed to go
				$TxtConnector = ("TEXT;" + $CSVFile)
				$CellRef = $worksheet.Range("A1")
				#Build, use and remove the text file connector
				$Connector = $worksheet.QueryTables.add($TxtConnector,$CellRef)
				$worksheet.QueryTables.item($Connector.name).TextFileCommaDelimiter = $True
				$worksheet.QueryTables.item($Connector.name).TextFileParseType  = 1
				$worksheet.QueryTables.item($Connector.name).Refresh()
				$worksheet.QueryTables.item($Connector.name).delete()
				$worksheet.UsedRange.EntireColumn.AutoFit()
				$listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, $null,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes,$null) 
				$listObject.Name = $TableTitle
				$listObject.TableStyle = "TableStyleLight2" # Style Cheat Sheet in French/English: http://msdn.microsoft.com/fr-fr/library/documentformat.openxml.spreadsheet.tablestyle.aspx 
				if ($isFreeze)
					{
						$worksheet.Activate();
						$worksheet.Application.ActiveWindow.SplitRow = 1;
						$worksheet.Application.ActiveWindow.FreezePanes = $isFreeze
					}
		<#
		$TmpWorkBook.SaveAs($excelFile,51) # http://msdn.microsoft.com/en-us/library/bb241279.aspx 
		$TmpWorkBook.Saved = $true
		$TmpWorkBook.Close() 
		ReleaseObject $TmpExcel
		return $excelFile
		#>
		
			}
	}
	
#---------------------------------------------------
# Function to release OBJ Excel
#---------------------------------------------------

function ReleaseObject([Object] $objExcel, [Object] $WorkBook, [Object] $WorkSheet)
	{
		$objExcel.DisplayAlerts = $false
		$objExcel.ActiveWorkBook.Close
		$objExcel.Quit()
		#[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$WorBook) | Out-Null
		#[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$WorkSheet) | Out-Null
		[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$objExcel) | Out-Null
		$objExcel = $null
		$WorkBook = $null
		$WorkSheet = $null
		[GC]::Collect()
		[GC]::WaitForPendingFinalizers()
	}
	
#---------------------------------------------------
# Function to Save Excel
#---------------------------------------------------

function SaveExcel([Object] $WorkBook, [String] $Path)
	{
		$ExcelReport = $Path + "\Report"
		$WorkBook.SaveAs($ExcelReport,51) # http://msdn.microsoft.com/en-us/library/bb241279.aspx 
		$WorkBook.Saved = $true
		$WorkBook.Close() 
	}

#---------------------------------------------------
# Function to Open Excel
#---------------------------------------------------

function Open-Excel([String] $Path)
	{
		#Open Excel 
		[Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'
		$Excel = new-object -com Excel.Application
		$Excel.visible = $false
		$WorkBook = $Excel.Workbooks.Open($Path)
		return $WorkBook		
	}

#---------------------------------------------------
# Function Create Table
#---------------------------------------------------

function Create-Table([Bool] $ShowTotals, [Object] $Data, [Object] $Header, [String] $Title, [String] $TableTitle, [Object] $Position, [Object] $WorkSheet, [Bool] $ShowHeaders)
	{
		$Cells = $WorkSheet.Cells
		$Row=$Position[0]
		$InitialRow = $Row
		$Col=$Position[1]
		$InitialCol = $Col
		if ($Header)
			{
				#insert column headings
				$Header | foreach `
					{
    					$cells.item($row,$col)=$_
    					$cells.item($row,$col).font.bold=$True
    					$Col++
					}
			}
		# Add table content
		foreach ($Key in $Data.Keys)
			{
	    		$Row++
	    		$Col = $InitialCol
	    		$cells.item($Row,$Col) = $Key
		
				$nbItems = $Data[$Key].Count
				for ( $i=0; $i -lt $nbItems; $i++ )
					{
					$Col++
	    			$cells.item($Row,$Col) = $Data[$Key][$i]
	    			$cells.item($Row,$Col).NumberFormat ="0"
					}		
			}
		# Apply Styles to table
		$Range = $WorkSheet.Range($WorkSheet.Cells.Item($InitialRow,$InitialCol),$WorkSheet.Cells.Item($Row,$Col))
		$listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $Range, $null,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes,$null) 
		$listObject.Name = $TableTitle
		$listObject.ShowTotals = $ShowTotals
		$listObject.ShowHeaders = $ShowHeaders
		$listObject.TableStyle = "TableStyleLight2" # Style Cheat Sheet in French/English: http://msdn.microsoft.com/fr-fr/library/documentformat.openxml.spreadsheet.tablestyle.aspx 
		
		# Sort data based on the 2nd column
		$SortRange = $WorkSheet.Range($WorkSheet.Cells.Item($InitialRow+1,$InitialCol+1).Address($False,$False)) # address: Convert cells position 1,1 -> A:1
		$WorkSheet.Sort.SortFields.Clear()
		[void]$WorkSheet.Sort.SortFields.Add($SortRange,0,1,0)
		$WorkSheet.Sort.SetRange($Range)
		$WorkSheet.Sort.Header = 1 # exclude header
		$WorkSheet.Sort.Orientation = 1
		$WorkSheet.Sort.Apply()
		
		# Apply Styles to Title
		$cells.item(1,$InitialCol) = $Title
		$RangeTitle = $WorkSheet.Range($WorkSheet.Cells.Item(1,$InitialCol),$WorkSheet.Cells.Item(1,$Col))
		$RangeTitle.MergeCells = $true
		$RangeTitle.Style = "Heading 3"
		# http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.constants.aspx
		$RangeTitle.HorizontalAlignment = -4108
		$RangeTitle.ColumnWidth = 20
		
	}


#---------------------------------------------------
# Create a chart in Excel
# http://www.alexwinner.com/articles/powershell/115-posh-excel-part4.html
#---------------------------------------------------
function Create-MyChart ([Object] $WorkSheet,$DataRange,$ChartRange, $ChartType,[String] $Title, [Bool] $HasDataTable, $Style, [Bool] $saveImage)
{
	# Add the chart
	$Chart = $WorkSheet.Shapes.AddChart().Chart
	$Chart.ChartType = $ChartType
	#$Chart | gm
	
	# Apply a specific style for each type
	if ( $ChartType -like "xlPie" )
	{	
		$Chart.ApplyLayout(2,$Chart.ChartType)
		$Chart.Legend.Position = -4107
		if ( $Title )
			{
				$Chart.HasTitle = $true
				$Chart.ChartTitle.Text = $Title
			}
		# http://msdn.microsoft.com/fr-fr/library/microsoft.office.interop.excel._chart.setsourcedata(v=office.11).aspx
		$Chart.SetSourceData($DataRange)
	}
	else
	{	
		$Chart.SetSourceData($DataRange,[Microsoft.Office.Interop.Excel.XLRowCol]::xlRows)
		
		# http://msdn.microsoft.com/en-us/library/office/bb241345(v=office.12).aspx
		$Chart.Legend.Position = -4107
		$Chart.ChartStyle = $Style
		$Chart.ApplyLayout(2,$Chart.ChartType)
		if ($HasDataTable)
			{
			$Chart.HasDataTable = $true
			$Chart.DataTable.HasBorderOutline = $true
			}
		
		$NbSeries = $Chart.SeriesCollection().Count
		
		# Define data labels
		for ( $i=1 ; $i -le $NbSeries; ++$i )
		{
			$Chart.SeriesCollection($i).HasDataLabels = $true
			$Chart.SeriesCollection($i).DataLabels(0).Position = 3
		}
		
		$Chart.HasAxis([Microsoft.Office.Interop.Excel.XlAxisType]::xlCategory) = $false
		$Chart.HasAxis([Microsoft.Office.Interop.Excel.XlAxisType]::xlValue) = $false
		if ( $Title )
			{
			$Chart.HasTitle = $true
			$Chart.ChartTitle.Text = $Title
			}
	}
	# Define the position of the chart
	$ChartObj = $Chart.Parent

	$ChartObj.Height = $ChartRange.Height
	$ChartObj.Width = $ChartRange.Width
	
	$ChartObj.Top = $ChartRange.Top
	$ChartObj.Left = $ChartRange.Left
	if ($saveImage)
		{
			$ImageFile = ($env:temp + "\" + ([System.Guid]::NewGuid()).ToString() + ".png")
			$Chart.Export($ImageFile)
			return $ImageFile
		}
}

#---------------------------------------------------
# Read Headers in Excel
# Return Hash Table of each cell
#---------------------------------------------------

Function Read-Headers ([Object] $WorkSheet)
	{   
    $Headers =@{}
    $column = 1
    Do {
        $Header = $Worksheet.cells.item(1,$column).text
        If ($Header) {
            $Headers.add($Header, $column)
            $column++
        }
    } until (!$Header)
    return $Headers
	}

#---------------------------------------------------
# Format Cells
# Add color format for each ACL value
#---------------------------------------------------

Function Add-CellColor([Object] $Excel, [String] $SheetName, [String]$ColumnName)
	{
		$WorkSheet = $Excel.WorkSheets.Item($SheetName)
		$Headers = Read-Headers $WorkSheet
		$count = $WorkSheet.Cells.Item(65536, 4).End(-4162)
		for($startRow=2;$startRow -le $count.row;$startRow++)
			{
			$Color = Get-Color $WorkSheet.Cells.Item($startRow, $Headers[$ColumnName]).Value()
			if ($Color)
				{
					$WorkSheet.Cells.Item($startRow, $Headers[$ColumnName]).Interior.ColorIndex = $Color
					$WorkSheet.Cells.Item($startRow, $Headers[$ColumnName]).Font.ColorIndex = 2
					$WorkSheet.Cells.Item($startRow, $Headers[$ColumnName]).Font.Bold = $true
				}
			}
	}

#---------------------------------------------------
# Format Cells
# Add Icon format for each ACL value
#---------------------------------------------------

Function Add-Icon([Object] $Excel, [String] $SheetName, [String]$ColumnName)
	{
		#Charts Variables
		$xlConditionValues=[Microsoft.Office.Interop.Excel.XLConditionValueTypes]
		$xlIconSet=[Microsoft.Office.Interop.Excel.XLIconSet]
		$xlDirection=[Microsoft.Office.Interop.Excel.XLDirection]
		$WorkSheet = $Excel.WorkSheets.Item($SheetName)
		$Headers = Read-Headers $WorkSheet
		#Add Icons
		$range = [char]($Headers["Risk"]+64)
		$start=$WorkSheet.range($range+"2")
		#get the last cell
		$Selection=$WorkSheet.Range($start,$start.End($xlDirection::xlDown))
		#add the icon set
		$Selection.FormatConditions.AddIconSetCondition() | Out-Null
		$Selection.FormatConditions.item($($Selection.FormatConditions.Count)).SetFirstPriority()
		$Selection.FormatConditions.item(1).ReverseOrder = $True
		$Selection.FormatConditions.item(1).ShowIconOnly = $True
		$Selection.FormatConditions.item(1).IconSet = $xlIconSet::xl3TrafficLights1
		$Selection.FormatConditions.item(1).IconCriteria.Item(2).Type=$xlConditionValues::xlConditionValueNumber
		$Selection.FormatConditions.item(1).IconCriteria.Item(2).Value=60
		$Selection.FormatConditions.item(1).IconCriteria.Item(2).Operator=7
		$Selection.FormatConditions.item(1).IconCriteria.Item(3).Type=$xlConditionValues::xlConditionValueNumber
		$Selection.FormatConditions.item(1).IconCriteria.Item(3).Value=90
		$Selection.FormatConditions.item(1).IconCriteria.Item(3).Operator=7
	}

#---------------------------------------------------
# Create Report Index
# Function to create report index with HyperLinks
#---------------------------------------------------

Function Create-Index([Object] $WorkBook, [Object] $Excel)
	{
		[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
		#Background Image
		$bytes = [System.Convert]::FromBase64String($IndexImage)
		$ImgFile = ($env:temp + "\" + ([System.Guid]::NewGuid()).ToString() + ".jpg")
		$p = New-Object IO.MemoryStream($bytes, 0, $bytes.length)
		$p.Write($bytes, 0, $bytes.length)
		$picture = [System.Drawing.Image]::FromStream($p, $true)
		$picture.Save($ImgFile)
		#Innotec Image
		$bytes = [System.Convert]::FromBase64String($InnotecImage)
		$InnotecFile = ($env:temp + "\" + ([System.Guid]::NewGuid()).ToString() + ".png")
		$p = New-Object IO.MemoryStream($bytes, 0, $bytes.length)
		$p.Write($bytes, 0, $bytes.length)
		$picture = [System.Drawing.Image]::FromStream($p, $true)
		$picture.Save($InnotecFile)
		#Main Report Index		
		$row = 6
		$col = 1
		$WorkSheet = $WorkBook.WorkSheets.Add()
		$WorkSheet.Name = "Index"
		$WorkSheet.Tab.ColorIndex = 5
		$WorkBook.WorkSheets | ForEach-Object `
			{
				$format = [char](64+$row)
				$r=$format+"$row"
				$range = $WorkSheet.Range($r)
				$v = $WorkSheet.Hyperlinks.Add($WorkSheet.Cells.Item($row,$col),"","'$($_.Name)'"+"!$($r)","","$($_.Name)")
				$row++
			}
		$CellRange = $WorkSheet.Range("A1:A20")
		$CellRange.Interior.ColorIndex = 49
		$CellRange.Font.ColorIndex = 2
		$CellRange.Font.Bold = $true
		$WorkSheet.columns.item("A").EntireColumn.AutoFit() | out-null
		$Excel.ActiveWindow.Displaygridlines = $false
		$v = $WorkSheet.Shapes.AddPicture($InnotecFile,1,0,0,0,112,64)
		$y = $WorkSheet.SetBackgroundPicture($ImgFile)
	}

#---------------------------------------------------
# Create About page
# Function to create About page with HyperLinks
#---------------------------------------------------
Function Create-About([Object] $WorkBook, [Object] $Excel)
	{
		#Main Report Index
		[Void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
		$WorkSheet = $WorkBook.WorkSheets.Add()
		$WorkSheet.Name = "About"
		$WorkSheet.Cells.Item(1,1).Value() = "Active Directory Attack Surface"
		$WorkSheet.Cells.Item(1,1).Font.Size = 25
		$WorkSheet.Cells.Item(1,1).Font.Bold = $true
		$WorkSheet.Cells.Item(17,4).Value() = "http://www.innotecsystem.com"
		$r = $WorkSheet.Range("D17")
		$v = $WorkSheet.Hyperlinks.Add($r,"http://www.innotecsystem.com")
		$WorkSheet.Cells.Item(17,4).Font.Size = 12
		$WorkSheet.Cells.Item(17,4).Font.Bold = $true
		$WorkSheet.Cells.Item(18,4).Value() = "http://windowstips.wordpress.com"
		$r = $WorkSheet.Range("D18")
		$v = $WorkSheet.Hyperlinks.Add($r,"http://windowstips.wordpress.com")
		$WorkSheet.Cells.Item(18,4).Font.Size = 12
		$WorkSheet.Cells.Item(18,4).Font.Bold = $true
		#$WorkSheet.Cells.Item(2,5).Value() = ""
		#Background Image
		$bytes = [System.Convert]::FromBase64String($EntelgyImage2)
		$EntelgyFile = ($env:temp + "\" + ([System.Guid]::NewGuid()).ToString() + ".png")
		$p = New-Object IO.MemoryStream($bytes, 0, $bytes.length)
		$p.Write($bytes, 0, $bytes.length)
		$picture = [System.Drawing.Image]::FromStream($p, $true)
		$picture.Save($EntelgyFile)
		$v= $WorkSheet.Shapes.AddPicture($EntelgyFile,0,1,0,10,616,290)
		$ShapeRange = $WorkSheet.Shapes.Range(1)
		#$ShapeRange.Left = -999996 #Right constant
		$ShapeRange.Left = 372
		$Excel.ActiveWindow.Displaygridlines = $false
		$CellRange = $WorkSheet.Range("A1:G20")
		$CellRange.Interior.ColorIndex = 14
		$CellRange.Font.ColorIndex = 2
		$row = 1
		$col = 1
	}