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
# Construct detailed PSObject
#---------------------------------------------------

function Set-PSObject ([Object] $Object, [String] $type)
	{
		$FinalObject = @()
		foreach ($Obj in $Object)
			{
				$NewObject = New-Object -TypeName PSObject -Property $Obj
            	$NewObject.PSObject.typenames.insert(0,[String]::Format($type))
				$FinalObject +=$NewObject
			}
		return $FinalObject
	}

#---------------------------------------------------
# Get Color of ACL
#---------------------------------------------------

Function Get-Color ([String] $Color)
	{
		try
			{
				return $ACLColors[$Color]
			}
		catch
			{
				Write-Host "ACL color not found" -ForegroundColor Red
			}
	}

#---------------------------------------------------
# Volumetry Organizational Unit
#---------------------------------------------------
function Get-OUVolumetry{
    Param (
            [parameter(Mandatory=$true, HelpMessage="ACL object")]
            [Object]$ACL,
            
            [parameter(Mandatory=$true, HelpMessage="Users object")]
            [Object]$Users 

    )
    Begin{
        #Start computers plugin
        Write-Host ("{0}: {1}" -f "OU task", "Generating organizational unit volumetry from $($Global:AllDomainData.Name)")`
        -ForegroundColor Magenta
    }
    Process{
            if($ACL){
                #Connect to domain
                $Connection = $AllDomainData.ADConnection
                $Searcher = New-Object System.DirectoryServices.DirectorySearcher -ArgumentList $Connection
                $Searcher.SearchScope = "subtree"
                $Searcher.SearchRoot = $Connection
                if($Searcher){
                    $Searcher.PageSize = 200
                }
		        $ListOU = @()
		        foreach ($result in $ACL){
				    $ou =[String]::Format($result.DN)
                    $ouname = [String]::Format($result.Name)               
				    $Searcher.filter = "(&(objectclass=organizationalunit)(distinguishedname=$ou))"
				    $tmpou = $Searcher.FindOne()
				    $data = $tmpou.GetDirectoryEntry()
				    $NumberOfUsers = 0
				    if ($data){
				        $TotalUsers = $data.Children | Where-Object {$_.schemaClassName -eq "user"} | Select-Object sAMAccountName
				        $InactiveUsers = 0
				        $ActiveUsers = 0
					    foreach ($u in $TotalUsers){
                            $match = $Users | Where-Object {$_.sAMAccountName -eq $u.sAMAccountNAme -and $_.isActive -eq $true}
                            if ($match){
                                $ActiveUsers++
                            }
                            else{
                                $InactiveUsers++
                            }                    
					    }
				    }
                    try{
                        $Count = [String]::Format($TotalUsers.Count)
                        $Active = [String]::Format($ActiveUsers)
                        $Inactive = [String]::Format($InactiveUsers)
                        $OUObject = New-Object -TypeName PSCustomObject
                        $OUObject | Add-Member -type NoteProperty -name Name -value $ouname
                        $OUObject | Add-Member -type NoteProperty -name DistinguishedName -value $ou
                        $OUObject | Add-Member -type NoteProperty -name NumberOfUsers -value $Count
                        $OUObject | Add-Member -type NoteProperty -name ActiveUsers -value $Active
                        $OUObject | Add-Member -type NoteProperty -name InactiveUsers -value $Inactive
                        #Add Metadata
        		        $OUObject.PSObject.typenames.insert(0,'Arsenal.AD.OuCount')
				        $ListOU +=$OUObject	
                        #Need to resolv filters
                    }
                    catch{
                        Write-Host "Get-OUVolumetry problem...$($_.Exception.Message) in $($NumberofUsers)" -ForegroundColor Red
                    }
			    }
	        }
        }
        End{
            #Create custom object for store data
            $OUObject = New-Object -TypeName PSCustomObject
            $OUObject | Add-Member -type NoteProperty -name Data -value $ListOU
            #Formatting Excel Data
            $Excelformatting = New-Object -TypeName PSCustomObject
            $Excelformatting | Add-Member -type NoteProperty -name Data -value $ListOU
            $Excelformatting | Add-Member -type NoteProperty -name TableName -value "OU Volumetry"
            $Excelformatting | Add-Member -type NoteProperty -name SheetName -value "OU Volumetry"
            $Excelformatting | Add-Member -type NoteProperty -name isFreeze -value $True
            $Excelformatting | Add-Member -type NoteProperty -name Type -value "CSVTable"

            #Add Excel formatting into psobject
            $OUObject | Add-Member -type NoteProperty -name Excelformat -Value $Excelformatting
            return $OUObject
        }
    }
	
function New-Report([String] $Path, [String] $Domain)
	{
		$target = "$Path\Reports"
		if (!(Test-Path -Path $target))
			{
				$tmpdir = New-Item -ItemType Directory -Path $target
				Write-Host "Folder Reports created in $target...." -ForegroundColor Yellow
			}
		$folder = "$target\" + ([System.Guid]::NewGuid()).ToString() + $Domain
		if (!(Test-Path -Path $folder))
			{
				try
					{
						$tmpdir = New-Item -ItemType Directory -Path $folder
						return $folder
					}
				catch
					{
						Write-Verbose "Failed to create new directory. Trying to generate new guid...."
						New-Report -Path $Path -Domain $Domain
					}
			}
				
	}



