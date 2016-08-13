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

#GroupTypes

#A few colors for format cells
$ACLColors = @{
			"CreateChild, DeleteChild" = 44; #Warning
			"CreateChild, Self, WriteProperty, ExtendedRight, Delete, GenericRead, WriteDacl, WriteOwner" = 44;#Warning
			"CreateChild, Self, WriteProperty, ExtendedRight, GenericRead, WriteDacl, WriteOwner" = 44;#Warning
            "CreateChild, DeleteChild, Self, WriteProperty, ExtendedRight, Delete, WriteDacl, WriteOwner" = 44;#Warning
			"GenericAll" = 9;#High
			"GenericRead" = 10; #Low
			"ReadProperty, WriteProperty, ExtendedRight" = 44;#Warning
			"ListChildren" = 10;#Low
			"ReadProperty" = 10;#Low
			"ReadProperty, WriteProperty" = 44;#Warning
			"ExtendedRight" = 44;#Warning
            "WriteProperty" = 44;#Warning
            "ListChildren, ReadProperty, ListObject" = 10; #Low
            "GenericRead, WriteDacl" = 10; #Low
            "ReadProperty, GenericExecute" = 44;#Warning
			"CreateChild, DeleteChild, Self, WriteProperty, ExtendedRight, Delete, GenericRead, WriteDacl, WriteOwner" = 44;#Warning
			"CreateChild, DeleteChild, Self, WriteProperty, ExtendedRight, GenericRead, WriteDacl, WriteOwner" = 44;#Warning
            "DeleteChild" = 66;#Warning
            "DeleteTree, Delete" = 66;#Warning
            "DeleteChild, DeleteTree, Delete" = 66;#Warning
            "Self, WriteProperty, GenericRead" = 66; #Warning
            "CreateChild, ReadProperty, GenericExecute" = 66; #Warning
            "CreateChild, ReadProperty, WriteProperty, GenericExecute" = 66; #Warning
            "ListChildren, ReadProperty, GenericWrite" = 66;#Warning
            "CreateChild, ListChildren, ReadProperty, GenericWrite" = 66;#Warning
            "CreateChild" = 10;#Low

			
			}

# Set Global variables
Set-Variable -Name InactiveDays -Value 180 -Scope Global
Set-Variable -Name CurrentDate -Value (Get-Date).ToFileTimeUTC() -Scope Global
Set-Variable -Name InactiveDate -Value (Get-Date).AddDays(-[int]$InactiveDays).ToFileTimeUtc() -Scope Global
Set-Variable -Name PasswordAge -Value 120 -Scope Global
Set-Variable -Name PasswordLastSet -Value (Get-Date).AddDays(-[int]$PasswordAge).ToFileTimeUtc() -Scope Global

					
#Count of roles
$Roles = @{
			SQLServer=0;
			TerminalServer=0;
			Exchange=0;
			SCOM=0;
			Cluster=0;
			GlobalCatalog=0;
			DNS=0;			
		}