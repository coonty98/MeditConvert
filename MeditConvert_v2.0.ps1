# This variable is manually set for testing purposes. In a
# production environment, this should be changed to pull the
# current directory or functionality should be added to pull
# the file path from a database to allow for easier modification.
$global:MfgrDir = '\\TYCOONCOMPUTER\Tyler Shared\Testing\'

$Hash_USFDI = @{
    "11" = "8"
    "12" = "7"
    "13" = "6"
    "14" = "5"
    "15" = "4"
    "16" = "3"
    "17" = "2"
    "18" = "1"
    "21" = "9"
    "22" = "10"
    "23" = "11"
    "24" = "12"
    "25" = "13"
    "26" = "14"
    "27" = "15"
    "28" = "16"
    "31" = "24"
    "32" = "23"
    "33" = "22"
    "34" = "21"
    "35" = "20"
    "36" = "19"
    "37" = "18"
    "38" = "17"
    "41" = "25"
    "42" = "26"
    "43" = "27"
    "44" = "28"
    "45" = "29"
    "46" = "30"
    "47" = "31"
    "48" = "32"
}

Function Set-ToothNumber{
	#########################################################################
    #   This function converts the tooth number from FDI to US, creates     #
	#	the XML ExportedObjects element for the margin, and creates the     #
	#	teeth nodes for each XML.                                           #
    #########################################################################

	Param([string]$casefolderpath)

    #This will change the margin line extensions to .pts
    Get-ChildItem -Path $global:MfgrDir"Medit Convert" -Recurse *..xyz | Rename-Item -NewName { $_.Name -replace '\..xyz$','.pts' }
    
	$margins = Get-ChildItem -Path $global:MfgrDir"Medit Convert\$casefolderpath\" -Recurse -Filter "*MarginLine*"
	if($margins.Count -eq 0){
        $margins = "Maxilla_MarginLine_18.pts"
    }
	ForEach($margin in $margins){
		$ToothNumber = $margin -replace "[^0-9]",''
		$USToothNumber = $Hash_USFDI[$ToothNumber]
		$ExObj_marginline += "<Object ObjectType=`"Line`" SubType=`"Margin`" AdaId=`"$USToothNumber`" FileName=`"$margin`"/>`r`n"
		#$TeethNode_v1021 += "<Tooth AdaId=`"$USToothNumber`" RestorationType=`"2`" ToothMaterial=`"60`" MarginDesignInt=`"`" MarginDesignExt=`"`" IncisalShade=`"A1`" MiddleShade=`"A1`" GingivalShade=`"A2`" DieShade=`"`" PreparationDesignBuccal=`"`" PreparationDesignLingual=`"`"/>`r`n"
		$TeethNode_v222324 += "<Tooth AdaId=`"$USToothNumber`" RestorationType=`"2`" Ditch=`"0`" ToothMaterial=`"60`" MarginDesignInt=`"`" MarginDesignExt=`"`" IncisalShade=`"A1`" MiddleShade=`"A1`" GingivalShade=`"A2`" DieShade=`"`" PreparationDesignBuccal=`"`" PreparationDesignLingual=`"`"/>`r`n"
		#$TeethNode_v50 += "<Tooth AdaId=`"$USToothNumber`" RestorationType=`"2`" Specification=`"1`" ImplantBasedRestorationType=`"-1`" AbutmentType=`"-1`" PonticDesign=`"-1`" Ditch=`"0`" ToothMaterial=`"2009`" MarginDesignInt=`"`" MarginDesignExt=`"`" IncisalShade=`"A1`" MiddleShade=`"A1`" GingivalShade=`"A2`" DieShade=`"`" PreparationDesignBuccal=`"`" PreparationDesignLingual=`"`" AbutmentMaterial=`"-1`"/>`r`n"
	}
	Return $USToothNumber, $ExObj_marginline, $TeethNode_v222324
}

Function Test-FolderName{
    #########################################################################
    #   This function attempts to isolate the pt name from the folder name. #
    #########################################################################

    Param([string]$casefolder)

    If($casefolder.Contains('Case') -and -not $casefolder.Contains(',') -and -not $casefolder.Contains('#') -and $foldername -notmatch ".*['].*['].*"){
        $parts = $casefolder -Split("'")
        $extractedname = $parts[0]
    }
    ElseIf($casefolder.Contains('#') -and -not $casefolder.Contains('Case')){
        $parts = $casefolder -Split(" ")
        $extractedname = $parts[0] + ' ' + $parts[1]
    }
	ElseIf($casefolder.Contains(',')){
        $parts = $casefolder.split(',')
        $firstname_withcase = $parts[1]
        $firstnamesplit = $firstname_withcase.split("'")
        $extractedname = $firstnamesplit[0].Trim() + ' ' + $parts[0]
    }
	ElseIf($casefolder.Contains('Case') -and $foldername.Contains('#')){
        $parts = $foldername.split("'")
        $extractedname = $parts[0]
	}
	ElseIf($casefolder -match ".*['].*['].*"){
        $parts = $casefolder.split("'")
        $extractedname = $parts[0].Trim() + ' ' + $parts[2].Trim()
    }
    Else{
        $extractedname = $casefolder
    }
    Return $extractedname
}

Function Set-ExportedObjects{
	#########################################################################
    #   This function sets ExportedObjects node in the XML files.           #
    #########################################################################

	Param([string]$subfolder)

	$lowerjaw = Get-ChildItem -Path $global:MfgrDir"Medit Convert\$subfolder\" -Filter 'Mandible_Base*'
    $upperjaw = Get-ChildItem -Path $global:MfgrDir"Medit Convert\$subfolder\" -Filter 'Maxilla_Base*'
    $lowerpreop = Get-ChildItem -Path $global:MfgrDir"Medit Convert\$subfolder\" -Filter 'Mandible_PreOP*'
    $upperpreop = Get-ChildItem -Path $global:MfgrDir"Medit Convert\$subfolder\" -Filter 'Maxilla_PreOP*'

	$ExObj_lowerjaw = "<Object ObjectType=`"Surface`" SubType=`"Jaw`" JawId=`"Lower`" FileName=`"$lowerjaw`" />"
	$ExObj_upperjaw = "<Object ObjectType=`"Surface`" SubType=`"Jaw`" JawId=`"Upper`" FileName=`"$upperjaw`" />"
	If($lowerpreop.Count -eq 0){
		$ExObj_lowerpreop = $lowerpreop
		$ExObj_lowerpreop_v2450 = $lowerpreop
	}
	Else{
		$ExObj_lowerpreop = "<Object ObjectType=`"Surface`" SubType=`"PreTreatment_Jaw`" FileName=`"$lowerpreop`"/>"
		$ExObj_lowerpreop_v2450 = "<Object ObjectType=`"Surface`" SubType=`"PreTreatment_Jaw`" JawId=`"Lower`" FileName=`"$lowerpreop`"/>"
	}
	If($upperpreop.Count -eq 0){
		$ExObj_upperpreop = $upperpreop
		$ExObj_upperpreop_v2450 = $upperpreop
	}
	Else{
		$ExObj_upperpreop = "<Object ObjectType=`"Surface`" SubType=`"PreTreatment_Jaw`" FileName=`"$upperpreop`"/>"
		$ExObj_upperpreop_v2450 = "<Object ObjectType=`"Surface`" SubType=`"PreTreatment_Jaw`" JawId=`"Upper`" FileName=`"$upperpreop`"/>"
	}
	
	Return $ExObj_lowerjaw, $ExObj_upperjaw, $ExObj_lowerpreop, $ExObj_upperpreop, $ExObj_lowerpreop_v2450, $ExObj_upperpreop_v2450
}



Function Convert-Medit{
    #########################################################################
    #   This function creates the XML files from the XML segments built.    #
    #########################################################################


ForEach($subfolder in Get-ChildItem -LiteralPath $global:MfgrDir"Medit Convert" -Directory){
    $foldername = $subfolder.Name
    $extractedname = Test-FolderName -casefolder: $foldername
	$USToothNumber, $ExObj_marginline, $TeethNode_v222324 = Set-ToothNumber -casefolderpath: $foldername
	$ExObj_lowerjaw, $ExObj_upperjaw, $ExObj_lowerpreop, $ExObj_upperpreop, $ExObj_lowerpreop_v2450, $ExObj_upperpreop_v2450 = Set-ExportedObjects -subfolder: $foldername
    $xmlv24 = New-XMLfiles -ptname: "$extractedname" -TeethNode_v222324: $TeethNode_v222324 -ExObj_lowerjaw: $ExObj_lowerjaw -ExObj_upperjaw: $ExObj_upperjaw -ExObj_lowerpreop: $ExObj_lowerpreop -ExObj_upperpreop: $ExObj_upperpreop -ExObj_lowerpreop_v2450: $ExObj_lowerpreop_v2450 -ExObj_upperpreop_v2450: $ExObj_upperpreop_v2450 -ExObj_marginline: $ExObj_marginline
    New-Item -Path $global:MfgrDir"Medit Convert\$foldername\iTero_Export_#999999999_v24.xml" -ItemType File -Value $xmlv24 | Out-Null
}
}

Function New-XMLfiles{
    #########################################################################
    #   This function builds all the XML files.                             #
    #########################################################################

    param([string]$ptname,
		  [string]$TeethNode_v222324,
		  [string]$ExObj_lowerjaw,
		  [string]$ExObj_upperjaw,
		  [string]$ExObj_lowerpreop,
		  [string]$ExObj_upperpreop,
		  [string]$ExObj_lowerpreop_v2450,
		  [string]$ExObj_upperpreop_v2450,
		  [string]$ExObj_marginline)

$xmlv24 = @"
<?xml version="1.0" encoding="utf-8"?>
<iTeroExport Version="2.4">
	<RxInfo>
		<CaseId>999999999</CaseId>
		<Modeled>True</Modeled>
		<FilePath/>
		<ExportTime>12/12/2024 17:27:25</ExportTime>
		<UtcOffset>0</UtcOffset>
		<LabTechnician/>
		<Patient>$ptname</Patient>
		<ChartNumber/>
		<Doctor>Null</Doctor>
		<DoctorLicense>Null</DoctorLicense>
		<DoctorID>Null</DoctorID>
		<LabName>Null</LabName>
		<LabID>479</LabID>
		<PracticeShipToAddress>Null</PracticeShipToAddress>
		<DueDate>12/23/2024</DueDate>
		<ScanRange>
			<UpperRange FromAdaId="11" ToAdaId="16"/>
			<LowerRange FromAdaId="17" ToAdaId="22"/>
		</ScanRange>
		<CaaSVersion>0.9.7.14</CaaSVersion>
		<iTeroVersion>2.15.0.35</iTeroVersion>
		<ExportScheme>3Shape</ExportScheme>
		<SchemeInfo>Exported files are compatible with 3Shape Dental System version 2008/1 or later</SchemeInfo>
		<ToothShadeSystem>VITA Lumin</ToothShadeSystem>
		<Teeth>
			$TeethNode_v222324
		</Teeth>
		<TeethBuccalTransforms/>
		<Bridges/>
		<Notes>Null</Notes>
		<Local_iDE ProductionCenterID="-1"/>
		<iDE_Order ProductionCenterID="0" CADCAMSystemTypeID="-1" SelectedProductionCenterName="" SelectedCadcamSystemName=""/>
		<CA_Order ProductionCenterID="0" SelectedProductionCenterName=""/>
	</RxInfo>
	<RxDefines>
		<RestorationTypes>
			<Type Id="2" Name="Crown"/>
		</RestorationTypes>
		<BridgeTypes>
			<Type Id="1" Name="Traditional"/>
		</BridgeTypes>
		<ToothInBridgeTypes>
			<Type Id="-1" Name="N/R"/>
		</ToothInBridgeTypes>
		<Materials>
			<Material Id="2009" Name="Ceramic: Zirconia"/>
		</Materials>
		<ImplantTypes>
			<Type Id="1" Name="TissueLevel NN" ScanBodyMfg="Straumann" ScanBodyMfgExtID="" ImplantManufacturerName="Straumann"/>
		</ImplantTypes>
	</RxDefines>
	<ExportedObjects>
		$ExObj_lowerjaw
		$ExObj_lowerpreop_v2450
		$ExObj_upperjaw
		$ExObj_upperpreop_v2450
		$ExObj_marginline
	</ExportedObjects>
	<ActualUsedTransforms>
		<Transformation>0.000000 -1.000000 0.000000 1.473943 0.000000 0.000000 1.000000 -6.370500 -1.000000 0.000000 0.000000 12.680452 0.000000 0.000000 0.000000 1.000000</Transformation>
	</ActualUsedTransforms>
	<ScanBodyRecognitionData/>
</iTeroExport>
"@

    #Return XML segments to the calling script
    Return $xmlv24
}

Convert-Medit