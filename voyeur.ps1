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

.EXAMPLE
	.\voyeur.ps1 -ExportToCSV:$true
	
.EXAMPLE
	.\voyeur.ps1 -ExportToEXCEL:$true
	
.EXAMPLE
	.\voyeur.ps1 -Domain "fqdndomain" -Credential "Domain\username"
	
.EXAMPLE
	.\voyeur.ps1 -Domain "fqdndomain" -ExportToEXCEL:$true
	
.EXAMPLE
	.\voyeur.ps1 -Domain "fqdndomain" -InactiveDays 180

.EXAMPLE
	.\voyeur.ps1 -Domain "fqdndomain" -PasswordAge 120
	
.PARAMETER Domain
	Collect data from the specified domain.

.PARAMETER InactiveDays
	The number of inactive days for a user account to define it as inactive (default 180). 
	
.PARAMETER PasswordAge
	The number of days to define a password as old (default 120). 
	
.PARAMETER UseSSL
	For SSL connection an valid username/password and domain passed is neccesary to passed
#>

[CmdletBinding()] 
param
(	
	[Parameter(Mandatory=$false)]
	[String] $DomainName,
	
	[Parameter(Mandatory=$false)]
	[int] $InactiveDays = 180,
	
	[Parameter(Mandatory=$false)]
	[Bool] $ExportToCSV = $true,
	
	[Parameter(Mandatory=$false)]
	[Bool] $ExportToEXCEL = $false,
	
	[Parameter(Mandatory=$false)]
	[int] $PasswordAge = 120,
	
	[Parameter(Mandatory=$false)]
	[String] $Username = $false,

    [Parameter(Mandatory=$false)]
	[String] $SearchRoot = "",
	
	[Parameter(Mandatory=$false, HelpMessage="User/Password and Domain required")]
	[String] $UseSSL = $false
)

#Region Import Modules
#---------------------------------------------------
# Import Modules
#---------------------------------------------------	
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent
. $ScriptPath\Common\Domain.ps1
. $ScriptPath\Common\Vars.ps1
. $ScriptPath\Common\Functions.ps1
. $ScriptPath\UserAccounts\Users.ps1
. $ScriptPath\Groups\Groups.ps1
. $ScriptPath\UserAccounts\HighPrivileges.ps1
. $ScriptPath\ACL\ACL.ps1
. $ScriptPath\Computers\Computers.ps1
. $ScriptPath\Computers\Services.ps1
. $ScriptPath\Dashboard\Status.ps1
. $ScriptPath\Reporting\ExcelReport.ps1
. $ScriptPath\Reporting\CSVReport.ps1
. $ScriptPath\GroupPolicy\Get-GPPPassword.ps1
. $ScriptPath\Images\Images.ps1

#EndRegion

#Region FunctionExcel
Function Create-ExcelReport
	{
		try
			{
				#Add Types
				Add-Type -AssemblyName Microsoft.Office.Interop.Excel
				$xlChart=[Microsoft.Office.Interop.Excel.XLChartType]
				##############Create Excel Report########################
				$objExcel = Create-Excel
				$Excel = $objExcel[1]
				$WorkBook = $objExcel[0]
				# Create About Page
				Create-About $WorkBook $Excel
				#Create WorkSheet of status of users
				$UsersExport = $FinalUsers | Sort-Object | Select-Object $UserProperties
                try
                    {
                        $ErrorActionPreference = "Continue"
				        Create-CSV2Table -Data $FinalComputers -WorkBook $WorkBook -Title "Computers" -TableTitle "Computers" -Excel $Excel -isFreeze $true | Out-Null
				        Create-CSV2Table -Data $UsersExport -WorkBook $WorkBook -Title "Users" -TableTitle "Users" -Excel $Excel -isFreeze $true | Out-Null
				        Create-CSV2Table -Data $ACL -WorkBook $WorkBook -Title "OrganizationalUnit" -TableTitle "ACL" -Excel $Excel -isFreeze $true | Out-Null
				        Create-CSV2Table -Data $VolumetryACL -WorkBook $WorkBook -Title "OrganizationalUnit Volumetry" -TableTitle "ACL Volumetry" -Excel $Excel -isFreeze $true | Out-Null
				        Create-CSV2Table -Data $AllRoles -WorkBook $WorkBook -Title "All Roles" -TableTitle "All Roles" -Excel $Excel | Out-Null
				        Create-CSV2Table -Data $FinalGroups -WorkBook $WorkBook -Title "All Groups" -TableTitle "All Groups" -Excel $Excel | Out-Null
				        Create-CSV2Table -Data $HighPrivileges -WorkBook $WorkBook -Title "HighPrivileges" -TableTitle "HighPrivileges" -Excel $Excel -isFreeze $true | Out-Null
				        Create-CSV2Table -Data $AdminSDHolder -WorkBook $WorkBook -Title "AdminSDHolder" -TableTitle "AdminSDHolder" -Excel $Excel | Out-Null
				        Create-CSV2Table -Data $GPPLeaks -WorkBook $WorkBook -Title "Group Policy Leaks" -TableTitle "GPPLeaks" -Excel $Excel | Out-Null
                    }
                catch
                    {
                        $ErrorActionPreference = "Continue"
                        Write-Host "$($_.Exception.Message)" -ForegroundColor Red
                    }
				
				#Add colors to ACL Column
                #Need to solve try catch
				Add-Icon -Excel $Excel -SheetName "OrganizationalUnit" -ColumnName "Rights" | Out-Null
				#Add colors to ACL Column
				Add-Icon -Excel $Excel -SheetName "AdminSDHolder" -ColumnName "Rights" | Out-Null
				
               
				##Organizational Unit Status
				$OUStatusActive = @{}
				$OUStatusInactive = @{}
				$c= 0
				foreach ($ou in $VolumetryACL)
					{
                        #$ou | fl
                        if ($ou.InactiveUsers -ne 0 -and $ou.ActiveUsers -ne 0)
                            { 
                                $OUStatusInactive.Add($ou.distinguishedname,@($ou.InactiveUsers))
						        $OUStatusActive.Add($ou.distinguishedname,@($ou.ActiveUsers))
                                $c++
                            }
					}
                
                $SheetName = "Organizational Unit Status"
				$WorkSheet = Create-WorkSheet $WorkBook $SheetName $Excel 
				$Title = "Organizational Unit Status. Active Users"
				Create-Table -ShowTotals $false -ShowHeaders $false -Data $OUStatusActive -Title $Title -TableTitle "OU ActiveStatus" -Position @(2,1) -Header $false -WorkSheet $WorkSheet | Out-Null
				$Title = "Organizational Unit Status. Inactive Users"
				Create-Table -ShowTotals $false -ShowHeaders $false -Data $OUStatusInactive -Title $Title -TableTitle "OU InactiveStatus" -Position @(2,10) -Header $false -WorkSheet $WorkSheet | Out-Null
          
				# Define the data range for the first chart
				$DataRange = $WorkSheet.Range($WorkSheet.Cells.Item(3,1),$WorkSheet.Cells.Item((2+$c),2))
				$ChartRange = $WorkSheet.Range($WorkSheet.Cells.Item(1,1),$WorkSheet.Cells.Item(18,7))
				# Call the function Create-MyChart to create the second chart
				$ChartTitle = "Active Users in Organizational Unit"
				Create-MyChart -WorkSheet $WorkSheet -DataRange $DataRange -ChartRange $ChartRange -ChartType $xlChart::xlColumnClustered -Title $ChartTitle -HasDataTable $false -Style 10 | Out-Null
				# Define the data range for the first chart
				$DataRange = $WorkSheet.Range($WorkSheet.Cells.Item(3,10),$WorkSheet.Cells.Item((2+$c),11))
				$ChartRange = $WorkSheet.Range($WorkSheet.Cells.Item(1,9),$WorkSheet.Cells.Item(18,15))
				# Call the function Create-MyChart to create the second chart
				$ChartTitle = "Inactive Users in Organizational Unit"
				Create-MyChart -WorkSheet $WorkSheet -DataRange $DataRange -ChartRange $ChartRange -ChartType $xlChart::xlColumnClustered -Title $ChartTitle -HasDataTable $false -Style 10 | Out-Null
                
				##DashBoard Computers & Servers $ Roles
				$ComputerStatus = @{}
                $Count = 0
				foreach ($key in $ComputersCount.keys)
					{
                        if ($ComputersCount[$($key)] -ne 0)
                            {
						        $ComputerStatus.Add($key, @($ComputersCount[$($key)]))
                                $Count ++
                            }
					}

				$SheetName = "Dashboard Computers"
				$WorkSheet = Create-WorkSheet $WorkBook $SheetName $Excel
				$Title = "Status of Computers"
				Create-Table -ShowTotals $true -ShowHeaders $false -Data $ComputerStatus -Title $Title -TableTitle "Computer Status" -Position @(2,1) -Header $false -WorkSheet $WorkSheet | Out-Null

				# Define the data range for the first chart
				$DataRange = $WorkSheet.Range($WorkSheet.Cells.Item(3,1),$WorkSheet.Cells.Item((3+$Count),2))
				$ChartRange = $WorkSheet.Range($WorkSheet.Cells.Item(1,1),$WorkSheet.Cells.Item(18,$Count))
				# Call the function Create-MyChart to create the second chart
				$ChartTitle = "Status of computers"
				Create-MyChart -WorkSheet $WorkSheet -HasDataTable $true -DataRange $DataRange -ChartRange $ChartRange -ChartType $xlChart::xlColumnClustered -Title $ChartTitle -Style 34 -saveImage $true | Out-Null
				#Roles
				#Table with roles count
				$Roles = $AllRoles | Group Service
				$RolStatus = @{}
				$c= 0
				foreach ($rol in $Roles)
					{
						$RolStatus.Add($rol.Name, @($rol.Count))
						$c++
					}
				$Title = "Status of Services"
				Create-Table -ShowTotals $true -ShowHeaders $false -Data $RolStatus -Title $Title -TableTitle "Rol Status" -Position @(2,10) -Header $false -WorkSheet $WorkSheet | Out-Null
				# Define the data range for the first chart
				$DataRange = $WorkSheet.Range($WorkSheet.Cells.Item(3,10),$WorkSheet.Cells.Item((3+$c),11))
				$ChartRange = $WorkSheet.Range($WorkSheet.Cells.Item(1,10),$WorkSheet.Cells.Item(18,15))
				# Call the function Create-MyChart to create the second chart
				$ChartTitle = "Discovered roles"
				Create-MyChart -WorkSheet $WorkSheet -DataRange $DataRange -ChartRange $ChartRange -ChartType $xlChart::xlColumnClustered -Title $ChartTitle -HasDataTable $true -Style 34 | Out-Null

				#Create WorkSheet with graph high Privileges
				$HighUsers = $HighPrivileges | group Group
				$HighChart = @{}
				foreach ($high in $HighUsers)
					{
						$HighChart.Add($high.Name,@($high.Count))
					}

				#Create WorkSheet and Table
				$SheetName = "High Privileges Chart"
				$WorkSheet = Create-WorkSheet $WorkBook $SheetName $Excel
				$Title = "High Privileges"
				Create-Table -ShowTotals $true -ShowHeaders $true -Data $HighChart -Title $Title -TableTitle "High Users" -Position @(2,1) -Header $false -WorkSheet $WorkSheet | Out-Null
				# Define the data range for the first chart
				$Count = ($HighChart).Count
				$DataRange = $WorkSheet.Range($WorkSheet.Cells.Item(3,1),$WorkSheet.Cells.Item((3+$Count),2))
				$ChartRange = $WorkSheet.Range($WorkSheet.Cells.Item(1,1),$WorkSheet.Cells.Item(18,($Count+8)))
				# Call the function Create-MyChart to create the second chart
				$ChartTitle = "High Privileges Accounts in Groups"
				Create-MyChart -WorkSheet $WorkSheet -DataRange $DataRange -ChartRange $ChartRange -ChartType $xlChart::xlColumnClustered -Title $ChartTitle -HasDataTable $true -Style 34 | Out-Null

				# Dashboard users
				$UsersVolumetry = @{ 
									Actives = @($UsersCount['isActive']['Count']); 
									Inactives = @($UsersCount['isInActive']['Count'])
									}
					
				$Title = "Volumetry of user accounts"
				$SheetName = "Dashboard Users"
				$WorkSheet = Create-WorkSheet $WorkBook $SheetName $Excel
				Create-Table -ShowTotals $false -ShowHeaders $false -Data $UsersVolumetry -Title $Title -TableTitle "User Volumetry" -Position @(2,1) -Header $false -WorkSheet $WorkSheet | Out-Null
				$UsersStatus = @{ 
							Locked = @($UsersCount['isActive']['isLocked'],$UsersCount['isInActive']['isLocked']);
							"Password Expired" = @($UsersCount['isActive']['isPwdExpired'],$UsersCount['isInActive']['isPwdExpired']);
							Expired = @($UsersCount['isActive']['isExpired'],$UsersCount['isInActive']['isExpired']);
							Disabled = @($UsersCount['isActive']['isDisabled'],$UsersCount['isInActive']['isDisabled']);
							}
				$Title = "Status of user accounts"
				Create-Table -ShowTotals $false -ShowHeaders $true -Data $UsersStatus -Title $Title -TableTitle "User Account Status" -Position @(3,4) -Header @('Status','Active Accounts','Inactive Accounts') -WorkSheet $WorkSheet | Out-Null

				$UsersConfig = @{ 
							"No delegated privilege (PTH Attack)" = @($UsersCount['isActive']['isPTH'],$UsersCount['isInActive']['isPTH']);
							"Password older than $PasswordAge days" = @($UsersCount['isActive']['isPwdOld'],$UsersCount['isInActive']['isPwdOld']);
							"Password not required" = @($UsersCount['isActive']['isPwdNotRequired'],$UsersCount['isInActive']['isPwdNotRequired']);
							"Password never expires" = @($UsersCount['isActive']['isPwdNeverExpires'],$UsersCount['isInActive']['isPwdNeverExpires']);
						}
					
					$Title = "Configuration of user accounts"
				Create-Table -ShowTotals $false -ShowHeaders $true -Data $UsersConfig -Title $Title -TableTitle "User Configuration" -Position @(3,8) -Header @('Options','Active Accounts','Inactive Accounts') -WorkSheet $WorkSheet | Out-Null

				# Define the data range for the first chart
				$DataRange = $WorkSheet.Range($WorkSheet.Cells.Item(3,1),$WorkSheet.Cells.Item(4,2))
				$ChartRange = $WorkSheet.Range($WorkSheet.Cells.Item(1,1),$WorkSheet.Cells.Item(18,2))

				$ChartTitle1 = "Volumetry of user accounts"

				# Call the function Create-MyChart to create the first chart
				Create-MyChart -WorkSheet $WorkSheet -DataRange $DataRange -ChartRange $ChartRange -ChartType $xlChart::xlPie -Title $ChartTitle1 -HasDataTable $true -Style 34 | Out-Null

				# Define the data range for the second chart
				$ChartTitle = "Status of active user accounts"
				$DataRange = $WorkSheet.Range($WorkSheet.Cells.Item(3,4),$WorkSheet.Cells.Item(7,5))
				$ChartRange = $WorkSheet.Range($WorkSheet.Cells.Item(1,4),$WorkSheet.Cells.Item(18,6))
					
				# Call the function Create-MyChart to create the second chart
				Create-MyChart -WorkSheet $WorkSheet -DataRange $DataRange -ChartRange $ChartRange -ChartType $xlChart::xlBarClustered -Title $ChartTitle -HasDataTable $true -Style 34 | Out-Null

				$ChartTitle = "Configuration of active user accounts"
					
				# Define the data range for the third chart
				$DataRange = $WorkSheet.Range($WorkSheet.Cells.Item(4,8),$WorkSheet.Cells.Item(7,9))
				$ChartRange = $WorkSheet.Range($WorkSheet.Cells.Item(1,8),$WorkSheet.Cells.Item(18,10))

				# Call the function Create-MyChart to create the third chart
				Create-MyChart -WorkSheet $WorkSheet -DataRange $DataRange -ChartRange $ChartRange -ChartType $xlChart::xlColumnClustered -Title $ChartTitle -HasDataTable $true -Style 34 | Out-Null

				# Delete Sheet1
				$Excel.WorkSheets.Item("Sheet1").Delete() | Out-Null

				# Create Report Index
				Create-Index $WorkBook $Excel | Out-Null
				
				# Save Excel
				SaveExcel $WorkBook $Report 
				
				# Release Object
				ReleaseObject $Excel $WorkBook $WorkSheet
			}
		catch
			{
				Write-Host "$($_.Exception.Message)" -ForegroundColor Red
			}
	}

#EndRegion

#Region Main
Try
	{
		Set-Variable DomainPassed -Value $DomainName -Scope Global
		if ($Username -and $DomainName)
			{
				Write-Host "Establishing connection to domain with credential passed..."
				Set-Variable credential -Value (Get-Credential -Credential $Username) -Scope Global
				# Get Domain Name
				Set-Variable -Name Domain -Value (Get-AuthDomain $DomainName $Username) -Scope Global
			}
		elseif ($DomainName)
			{
				Write-Host "Establishing connection to the domain passed..."
				# Get Domain Name
				Set-Variable -Name DomainWPass -Value (Get-Domain $DomainName) -Scope Global
			}
		else
			{
				Set-Variable -Name CurrentDomain -Value (Get-CurrentDomain) -Scope Global
			}

        Set-Variable -Name Forest -Value (Get-CurrentForest) -Scope Global 
        Set-Variable -Name DomainSID -Value (Get-DomainSID) -Scope Global
        Set-Variable -Name RootDomainSID -Value (Get-RootDomainSID) -Scope Global
        
		# Get All Users in Domain
		$UserProperties = @("mail","accountExpires","distinguishedName","sAMAccountName","givenName",`
							"sn","description","userAccountControl","pwdLastSet","whenCreated",`
							"whenChanged","lastLogonTimestamp")
		$Users = Get-UsersInfo -Properties $UserProperties -Filter $SearchRoot
	    $Users = UACBitResolv $Users
		$Users = Get-MsDSUACResolv -Users $Users 
		$Users = Get-UserStatus $Users
		$FinalUsers = Set-PSObject $Users "Arsenal.AD.Users"
		$u = $FinalUsers[1]
		$tmp = [String]::Format(($u.distinguishedName.Split(",") | ? {$_ -like "DC=*"}))
		$NetBiosDomain = ($tmp.Replace("DC=","")).Replace(" ",".")
		Set-Variable -Name NetBios -Value $NetBiosDomain -Scope Global

		#Get All Groups in Domain
		$GroupProperties = @("name","distinguishedName","sAMAccountName","groupType","whenCreated","whenChanged")
		$Groups = Get-GroupInfo -Properties $GroupProperties -Filter $SearchRoot
		$FinalGroups = Set-PSObject $Groups "Arsenal.AD.Groups"			
		#Search for plain text passwords
		$GPPLeaks = Get-GPPPassword 
		#Search for Know Roles
		$AllRoles = Get-Roles
		# Get All Computers in Domain
            
		$Properties = @("distinguishedName","sAMAccountName","name","description"`
						,"operatingSystem","operatingSystemServicePack",`
						"operatingSystemVersion","whenCreated","whenChanged","lastLogon", "pwdLastSet")

		$Computers = Get-ComputersInfo -Properties $Properties -Filter $SearchRoot
        
		Get-ComputerStatus $Computers
		$FinalComputers = Set-PSObject $Computers "Arsenal.AD.Computers"
		# Get users with High Privileges in Domain
        
		$HighPrivileges = Get-HighPrivileges
		# Get All OU with ACL Report
        
		$ACL = GetACL ("objectclass=organizationalunit") -full $true -Filter $SearchRoot
		# Get All OU 
		$ACLOnly = GetACL ("objectclass=organizationalunit") -full $false -Filter $SearchRoot
		$VolumetryACL = Get-OUVolumetry $ACLOnly $FinalUsers 
		# Get ACL Report of AdminSDHolder object
        
		$AdminSDHolder = GetACL ("(&(name=AdminSDHolder)(objectCategory=Container))") -full $true

		#Create Report folder
		if ($ExportToCSV -eq $true -or $ExportToEXCEL -eq $true)
			{
				Write-Host "Create report folder...." -ForegroundColor Green
				$ReportPath = New-Report $ScriptPath $Domain.name
				Set-Variable -Name Report -Value $ReportPath -Scope Global
				Write-Host "Report folder created in $Report..." -ForegroundColor Green
				
			}
		#Export Data to CSV
		if ($ExportToCSV -eq $true)
			{
				Write-Host "Export data to CSV format...." -ForegroundColor Yellow
				#$ReportPath = New-Report $ScriptPath $Domain.name
				#Set-Variable -Name Report -Value $ReportPath -Scope Global
				try
					{
						Save-CSVData $FinalUsers $Report "All users"
						Save-CSVData $FinalGroups $Report "All Groups"
						Save-CSVData $FinalComputers $Report "All Computers"
						Save-CSVData $GPPLeaks $Report "Data Leakage"
						Save-CSVData $AllRoles $Report "All roles"
						Save-CSVData $AdminSDHolder $Report "AdminSDHolder ACL"
						Save-CSVData $ACL $Report "All Organizational Unit"
                        Save-CSVData $HighPrivileges $Report "High Privileges"
					}
				catch
					{
						Write-Host "$($_.Exception.Message)" -ForegroundColor Red
					}
			}#>
		if ($ExportToExcel -eq $true)
			{
				if (!$Global:Report)
					{
						Write-Host "No Report folder created..." -ForegroundColor Red
						$ReportPath = New-Report $ScriptPath $Domain.name
						Set-Variable -Name Report -Value $ReportPath -Scope Global
					}
				Write-Host "Export data to EXCEL format...." -ForegroundColor Yellow
				Create-ExcelReport
			}
            
	}
Catch
	{
		Write-Host "Voyeur problem...$($_.Exception.Message)" -ForegroundColor Red
	}

#Region Release Vars
$Report = $null