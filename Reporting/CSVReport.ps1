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
# Function to export object through CSV data
#---------------------------------------------------
Function Save-CSVData([Object] $Data, [String] $path, [String] $objectType)
	{
		Write-Host "Trying to save $objectType data...." -ForegroundColor Yellow
		try
			{
				if ($Data -ne $null)
					{
						$FinalPath = ($path + "\" + ([System.Guid]::NewGuid()).ToString() + "$objectType.csv")
						$Data | Export-Csv -path $FinalPath -noTypeInformation
						Write-Host "$objectType saved in $path successfully!" -ForegroundColor Green
					}
			}
		catch
			{
				Write-Host "$($_.Exception.Message)" -ForegroundColor Red
			}
	}