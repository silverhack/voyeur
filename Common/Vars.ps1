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
$GroupType = @{
			"-2147483646" = "Global Security Group";
			"-2147483644" = "Local Security Group";
			"-2147483643" = "BuiltIn Group";
			"-2147483640" = "Universal Security Group";
			"2" = "Global Distribution Group";
			"4" = "Local Distribution Group";
			"8" = "Universal Distribution Group";
			}

#Common Service Principal Names
$Services = @(@{Service="SQLServer";SPN="MSSQLSvc"},@{Service="TerminalServer";SPN="TERMSRV"},
					@{Service="Exchange";SPN="IMAP4"},@{Service="Exchange";SPN="IMAP"},
					@{Service="Exchange";SPN="SMTPSVC"},@{Service="SCOM";SPN="AdtServer"},
					@{Service="SCOM";SPN="MSOMHSvc"},@{Service="Cluster";SPN="MSServerCluster"},
					@{Service="Cluster";SPN="MSServerClusterMgmtAPI"};@{Service="GlobalCatalog";SPN="GC"},
					@{Service="DNS";SPN="DNS"},@{Service="Exchange";SPN="exchangeAB"}, 
                    @{Service ="WebServer";SPN="tapinego"},@{Service ="WinRemoteAdministration";SPN="WSMAN"},
                    @{Service ="ADAM";SPN="E3514235-4B06-11D1-AB04-00C04FC2DCD2-ADAM"},
                    @{Service ="Exchange";SPN="exchangeMDB"} )

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
#A few colors for format cells
$ACLIcons = @{
			"CreateChild, DeleteChild" = 66; #Warning
			"CreateChild, Self, WriteProperty, ExtendedRight, Delete, GenericRead, WriteDacl, WriteOwner" = 66;#Warning
            "CreateChild, DeleteChild, Self, WriteProperty, ExtendedRight, Delete, WriteDacl, WriteOwner" = 66;#Warning
			"CreateChild, Self, WriteProperty, ExtendedRight, GenericRead, WriteDacl, WriteOwner" = 66;#Warning
            "ListChildren, ReadProperty, ListObject" = 10; #Low
			"GenericAll" = 99;#High
			"GenericRead" = 10; #Low
			"ReadProperty, WriteProperty, ExtendedRight" = 66;#Warning
			"ListChildren" = 10;#Low
			"ReadProperty" = 10;#Low
			"ReadProperty, WriteProperty" = 66;#Warning
            "ReadProperty, GenericExecute" = 66;#Warning
            "GenericRead, WriteDacl" = 10; #Low
			"ExtendedRight" = 66;#Warning
            "WriteProperty" = 66;#Warning
            "DeleteChild" = 66;#Warning
			"CreateChild, DeleteChild, Self, WriteProperty, ExtendedRight, Delete, GenericRead, WriteDacl, WriteOwner" = 66;#Warning
			"CreateChild, DeleteChild, Self, WriteProperty, ExtendedRight, GenericRead, WriteDacl, WriteOwner" = 66;#Warning
			"DeleteTree, Delete" = 66;#Warning
            "DeleteChild, DeleteTree, Delete" = 66;#Warning
            "CreateChild, ReadProperty, GenericExecute" = 66; #Warning
            "Self, WriteProperty, GenericRead" = 66; #Warning
            "CreateChild, ReadProperty, WriteProperty, GenericExecute" = 66; #Warning
            "ListChildren, ReadProperty, GenericWrite" = 66;#Warning
            "CreateChild, ListChildren, ReadProperty, GenericWrite" = 66;#Warning
            "CreateChild" = 10;#Low
			}

# Groups with High privileges
$PrivGroups = @("Administrators";"Domain Admins";"Backup Operators";`
						"Server Operators";"Account Operators";"Enterprise Admins";"Schema Admins",`
						"Print Operators","Replicator","Read-only Domain Controllers")

# Groups with High privileges
$PrivGroupsESP = @("Administradores";"Admins. del dominio";"Operadores de copia de seguridad";`
						"Opers. de servidores";"Opers. de cuentas";"Administradores de empresas";"Administradores de esquema",`
						"Opers. de impresión","Duplicadores","Controladores de dominio de sólo lectura")

# Count user info
$UsersCount = @{
				isActive = @{
				Count = 0;
				isExpired = 0;
				isDisabled = 0;
				isPwdOld = 0;
				isLocked = 0;
				isPwdNeverExpires = 0;
				isPwdNotRequired = 0;
				isPwdExpired = 0
				isPTH = 0;
				}
				isInactive = @{
				Count = 0;
				isExpired = 0;
				isDisabled = 0;
				isPwdOld = 0;
				isLocked = 0;
				isPwdNeverExpires = 0;
				isPwdNotRequired = 0;
				isPwdExpired = 0
				isPTH = 0;
				}
		}

# Count computers info
$ComputersCount = @{
				WindowsXPSP1 = 0;
				WindowsXPSP2 = 0;
				WindowsXPSP3 = 0;
				WindowsServer2008SP1 = 0;
                WindowsServer2008SP2 = 0;
                WindowsServer2008R2SP1 = 0;
                WindowsServer2008R2STD = 0;
                WindowsServer2008R2STDSP1 = 0;
                WindowsServer2012STDStorage = 0;
                WindowsServer2008STDSP2 = 0;
                WindowsServer2008STDSP1 = 0;
                WindowsServer2008STD = 0;
				Windows7 = 0;
				Windows8 = 0;
                Windows8Pro = 0;
                Windows81Pro = 0;
                Windows7EnterpriseSP1 = 0;
                Windows7UltimateSP1 = 0;
                Windows7Ultimate = 0;
                Windows7ProfessionalSP1 = 0;
                Windows7Professional = 0;
				WindowsServer2003 = 0;
                WindowsServer2003SP1 = 0;
                WindowsServer2003SP2 = 0;
				WindowsServer2012 = 0;
                Windows2000SP4 = 0;
                Samba = 0;
                EMCFileServer = 0;
                LikeWiseOpen = 0;
				}
		
# Set Global variables
Set-Variable -Name InactiveDays -Value 180 -Scope Global
Set-Variable -Name CurrentDate -Value (Get-Date).ToFileTimeUTC() -Scope Global
Set-Variable -Name InactiveDate -Value (Get-Date).AddDays(-[int]$InactiveDays).ToFileTimeUtc() -Scope Global
Set-Variable -Name PasswordAge -Value 120 -Scope Global
Set-Variable -Name PasswordLastSet -Value (Get-Date).AddDays(-[int]$PasswordAge).ToFileTimeUtc() -Scope Global

#Bits of UserAccountControl
$UAC_USER_FLAG_ENUM = @{
						isLocked = 16;
						isDisabled = 2;
						isPwdNeverExpires = 65536;
						isNoDelegation = 1048576;
						isPwdNoRequired = 32;
						isPwdExpired = 8388608;}
						
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