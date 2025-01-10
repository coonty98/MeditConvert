$parentPath = $PSScriptroot

###Convert margin files from xyz to pts.
Get-ChildItem -Path $parentPath -Recurse *..xyz | Rename-Item -NewName { $_.Name -replace '\..xyz$','.pts' }

###Hash table converting FDI notation to US.
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

###Get path of master folder and list of Medit Files.
$combinePath = Join-Path -path $parentPath -ChildPath "Medit_Export_Master\*"
$MeditFiles = Get-ChildItem -Path $parentPath -Exclude 'Medit_Export_Master', 'MeditConvert.ps1'

###Loop through each Medit folder in the parent folder.
foreach ($MeditFile in $MeditFiles){
    ###Copy XML files from Master folder to each Medit folder.
    ###Assign path of XML files to variables.
    Copy-Item -Path $combinePath -Destination $MeditFile
    $pathv10 = Join-Path -Path $MeditFile -ChildPath "iTero_Export_#999999999_v10.xml"
    $pathv21 = Join-Path -Path $MeditFile -ChildPath "iTero_Export_#999999999_v21.xml"
    $pathv22 = Join-Path -Path $MeditFile -ChildPath "iTero_Export_#999999999_v22.xml"
    $pathv23 = Join-Path -Path $MeditFile -ChildPath "iTero_Export_#999999999_v23.xml"
    $pathv24 = Join-Path -Path $MeditFile -ChildPath "iTero_Export_#999999999_v24.xml"
    $pathv50 = Join-Path -Path $MeditFile -ChildPath "iTero_Export_#999999999_v50.xml"
    $Patient = $MeditFile.Name
    $LowerJawScan = 'Mandible_Base.ply'
    $UpperJawScan = 'Maxilla_Base.ply'

    ###Update patient in XML file using Medit Folder name.
    $xmlv10 = [xml](Get-Content $pathv10)
    $xmlv10.iTeroExport.RxInfo.Patient = $Patient
    $xmlv21 = [xml](Get-Content $pathv21)
    $xmlv21.iTeroExport.RxInfo.Patient = $Patient
    $xmlv22 = [xml](Get-Content $pathv22)
    $xmlv22.iTeroExport.RxInfo.Patient = $Patient
    $xmlv23 = [xml](Get-Content $pathv23)
    $xmlv23.iTeroExport.RxInfo.Patient = $Patient
    $xmlv24 = [xml](Get-Content $pathv24)
    $xmlv24.iTeroExport.RxInfo.Patient = $Patient
    $xmlv50 = [xml](Get-Content $pathv50)
    $xmlv50.iTeroExport.RxInfo.Patient = $Patient

    ###List all files with MarginLine in the name.
    $MarginFiles = Get-ChildItem -Path $MeditFile -Recurse -filter "*MarginLine*" | Select-Object Name

    ###Add the TEETH node.
    $newnodev10 = $xmlv10.iTeroExport.RxInfo.AppendChild($xmlv10.CreateElement("Teeth"))
    $newnodev21 = $xmlv21.iTeroExport.RxInfo.AppendChild($xmlv21.CreateElement("Teeth"))
    $newnodev22 = $xmlv22.iTeroExport.RxInfo.AppendChild($xmlv22.CreateElement("Teeth"))
    $newnodev23 = $xmlv23.iTeroExport.RxInfo.AppendChild($xmlv23.CreateElement("Teeth"))
    $newnodev24 = $xmlv24.iTeroExport.RxInfo.AppendChild($xmlv24.CreateElement("Teeth"))
    $newnodev50 = $xmlv50.iTeroExport.RxInfo.AppendChild($xmlv50.CreateElement("Teeth"))

    ###Assign arbitrary value if no margin file exists.
    if($MarginFiles.Count -eq 0){
        $MarginFiles = "Maxilla_MarginLine_18.pts"
    }else{
        $MarginFiles
    }

    ###Update XML files: Add tooth info.
    ###Obtain tooth number from margin file name by extracting all numbers from filename.
    ###Use hash table to convert tooth number.
    foreach ($MarginFile in $MarginFiles){
        $ToothNumber = $MarginFile -replace "[^0-9]",''
        $USToothNumber = $Hash_USFDI[$ToothNumber]

        $newelementv10 = $newnodev10.AppendChild($xmlv10.CreateElement("Tooth"))
        $newelementv10.SetAttribute("AdaId",$USToothNumber)
        $newelementv10.SetAttribute("RestorationType","2")
        $newelementv10.SetAttribute("ToothMaterial","60")
        $newelementv10.SetAttribute("MarginDesignInt","")
        $newelementv10.SetAttribute("MarginDesignExt","")
        $newelementv10.SetAttribute("IncisalShade","A2")
        $newelementv10.SetAttribute("MiddleShade","A2")
        $newelementv10.SetAttribute("GingivalShade","A3")
        $newelementv10.SetAttribute("DieShade","")
        $newelementv10.SetAttribute("PreparationDesignBuccal","")
        $newelementv10.SetAttribute("PreparationDesignLingual","")
        $newelementv21 = $newnodev21.AppendChild($xmlv21.CreateElement("Tooth"))
        $newelementv21.SetAttribute("AdaId",$USToothNumber)
        $newelementv21.SetAttribute("RestorationType","2")
        $newelementv21.SetAttribute("ToothMaterial","60")
        $newelementv21.SetAttribute("MarginDesignInt","")
        $newelementv21.SetAttribute("MarginDesignExt","")
        $newelementv21.SetAttribute("IncisalShade","A2")
        $newelementv21.SetAttribute("MiddleShade","A2")
        $newelementv21.SetAttribute("GingivalShade","A3")
        $newelementv21.SetAttribute("DieShade","")
        $newelementv21.SetAttribute("PreparationDesignBuccal","")
        $newelementv21.SetAttribute("PreparationDesignLingual","")
        $newelementv22 = $newnodev22.AppendChild($xmlv22.CreateElement("Tooth"))
        $newelementv22.SetAttribute("AdaId",$USToothNumber)
        $newelementv22.SetAttribute("RestorationType","2")
        $newelementv22.SetAttribute("Ditch","0")
        $newelementv22.SetAttribute("ToothMaterial","60")
        $newelementv22.SetAttribute("MarginDesignInt","")
        $newelementv22.SetAttribute("MarginDesignExt","")
        $newelementv22.SetAttribute("IncisalShade","A2")
        $newelementv22.SetAttribute("MiddleShade","A2")
        $newelementv22.SetAttribute("GingivalShade","A3")
        $newelementv22.SetAttribute("DieShade","")
        $newelementv22.SetAttribute("PreparationDesignBuccal","")
        $newelementv22.SetAttribute("PreparationDesignLingual","")
        $newelementv23 = $newnodev23.AppendChild($xmlv23.CreateElement("Tooth"))
        $newelementv23.SetAttribute("AdaId",$USToothNumber)
        $newelementv23.SetAttribute("RestorationType","2")
        $newelementv23.SetAttribute("Ditch","0")
        $newelementv23.SetAttribute("ToothMaterial","60")
        $newelementv23.SetAttribute("MarginDesignInt","")
        $newelementv23.SetAttribute("MarginDesignExt","")
        $newelementv23.SetAttribute("IncisalShade","A2")
        $newelementv23.SetAttribute("MiddleShade","A2")
        $newelementv23.SetAttribute("GingivalShade","A3")
        $newelementv23.SetAttribute("DieShade","")
        $newelementv23.SetAttribute("PreparationDesignBuccal","")
        $newelementv23.SetAttribute("PreparationDesignLingual","")
        $newelementv24 = $newnodev24.AppendChild($xmlv24.CreateElement("Tooth"))
        $newelementv24.SetAttribute("AdaId",$USToothNumber)
        $newelementv24.SetAttribute("RestorationType","2")
        $newelementv24.SetAttribute("Ditch","0")
        $newelementv24.SetAttribute("ToothMaterial","60")
        $newelementv24.SetAttribute("MarginDesignInt","")
        $newelementv24.SetAttribute("MarginDesignExt","")
        $newelementv24.SetAttribute("IncisalShade","A2")
        $newelementv24.SetAttribute("MiddleShade","A2")
        $newelementv24.SetAttribute("GingivalShade","A3")
        $newelementv24.SetAttribute("DieShade","")
        $newelementv24.SetAttribute("PreparationDesignBuccal","")
        $newelementv24.SetAttribute("PreparationDesignLingual","")
        $newelementv50 = $newnodev50.AppendChild($xmlv50.CreateElement("Tooth"))
        $newelementv50.SetAttribute("AdaId",$USToothNumber)
        $newelementv50.SetAttribute("RestorationType","2")
        $newelementv50.SetAttribute("Specification","1")
        $newelementv50.SetAttribute("ImplantBasedRestorationType","-1")
        $newelementv50.SetAttribute("AbutmentType","-1")
        $newelementv50.SetAttribute("PonticDesign","-1")
        $newelementv50.SetAttribute("Ditch","0")
        $newelementv50.SetAttribute("ToothMaterial","2010")
        $newelementv50.SetAttribute("MarginDesignInt","")
        $newelementv50.SetAttribute("MarginDesignExt","")
        $newelementv50.SetAttribute("IncisalShade","")
        $newelementv50.SetAttribute("MiddleShade","A2")
        $newelementv50.SetAttribute("GingivalShade","A3")
        $newelementv50.SetAttribute("DieShade","")
        $newelementv50.SetAttribute("PreparationDesignBuccal","")
        $newelementv50.SetAttribute("PreparationDesignLingual","")
        $newelementv50.SetAttribute("AbutmentMaterial","-1")

        ###List all files within Medit folder.
        $MeditFileChildren = Get-ChildItem -Path $MeditFile -Recurse

        ###If there are preop files, add elements with file info.
        if($MeditFileChildren -like '*Maxilla_PreOP*'){
            $MaxillaPreopElementv10 = $xmlv10.iTeroExport.ExportedObjects.AppendChild($xmlv10.CreateElement("Object"))
            $MaxillaPreopElementv10.SetAttribute("ObjectType","Surface")
            $MaxillaPreopElementv10.SetAttribute("SubType","PreTreatment_Jaw")
            $MaxillaPreopElementv10.SetAttribute("JawId","Upper")
            $MaxillaPreopElementv10.SetAttribute("Index",$USToothNumber)
            $MaxillaPreopElementv10.SetAttribute("FileName","Maxilla_PreOP.ply")
            $MaxillaPreopElementv21 = $xmlv21.iTeroExport.ExportedObjects.AppendChild($xmlv21.CreateElement("Object"))
            $MaxillaPreopElementv21.SetAttribute("ObjectType","Surface")
            $MaxillaPreopElementv21.SetAttribute("SubType","PreTreatment_Jaw")
            $MaxillaPreopElementv21.SetAttribute("JawId","Upper")
            $MaxillaPreopElementv21.SetAttribute("Index",$USToothNumber)
            $MaxillaPreopElementv21.SetAttribute("FileName","Maxilla_PreOP.ply")
            $MaxillaPreopElementv22 = $xmlv22.iTeroExport.ExportedObjects.AppendChild($xmlv22.CreateElement("Object"))
            $MaxillaPreopElementv22.SetAttribute("ObjectType","Surface")
            $MaxillaPreopElementv22.SetAttribute("SubType","PreTreatment_Jaw")
            $MaxillaPreopElementv22.SetAttribute("JawId","Upper")
            $MaxillaPreopElementv22.SetAttribute("Index",$USToothNumber)
            $MaxillaPreopElementv22.SetAttribute("FileName","Maxilla_PreOP.ply")
            $MaxillaPreopElementv23 = $xmlv23.iTeroExport.ExportedObjects.AppendChild($xmlv23.CreateElement("Object"))
            $MaxillaPreopElementv23.SetAttribute("ObjectType","Surface")
            $MaxillaPreopElementv23.SetAttribute("SubType","PreTreatment_Jaw")
            $MaxillaPreopElementv23.SetAttribute("JawId","Upper")
            $MaxillaPreopElementv23.SetAttribute("Index",$USToothNumber)
            $MaxillaPreopElementv23.SetAttribute("FileName","Maxilla_PreOP.ply")
            $MaxillaPreopElementv24 = $xmlv24.iTeroExport.ExportedObjects.AppendChild($xmlv24.CreateElement("Object"))
            $MaxillaPreopElementv24.SetAttribute("ObjectType","Surface")
            $MaxillaPreopElementv24.SetAttribute("SubType","PreTreatment_Jaw")
            $MaxillaPreopElementv24.SetAttribute("JawId","Upper")
            $MaxillaPreopElementv24.SetAttribute("Index",$USToothNumber)
            $MaxillaPreopElementv24.SetAttribute("FileName","Maxilla_PreOP.ply")
            $MaxillaPreopElementv50 = $xmlv50.iTeroExport.ExportedObjects.AppendChild($xmlv50.CreateElement("Object"))
            $MaxillaPreopElementv50.SetAttribute("ObjectType","Surface")
            $MaxillaPreopElementv50.SetAttribute("SubType","PreTreatment_Jaw")
            $MaxillaPreopElementv50.SetAttribute("JawId","Upper")
            $MaxillaPreopElementv50.SetAttribute("Index",$USToothNumber)
            $MaxillaPreopElementv50.SetAttribute("FileName","Maxilla_PreOP.ply")
        }

        if($MeditFileChildren -like '*Mandible_PreOP*'){
            $MandiblePreopElementv10 = $xmlv10.iTeroExport.ExportedObjects.AppendChild($xmlv10.CreateElement("Object"))
            $MandiblePreopElementv10.SetAttribute("ObjectType","Surface")
            $MandiblePreopElementv10.SetAttribute("SubType","PreTreatment_Jaw")
            $MandiblePreopElementv10.SetAttribute("JawId","Lower")
            $MandiblePreopElementv10.SetAttribute("Index",$USToothNumber)
            $MandiblePreopElementv10.SetAttribute("FileName","Mandible_PreOP.ply")
            $MandiblePreopElementv21 = $xmlv21.iTeroExport.ExportedObjects.AppendChild($xmlv21.CreateElement("Object"))
            $MandiblePreopElementv21.SetAttribute("ObjectType","Surface")
            $MandiblePreopElementv21.SetAttribute("SubType","PreTreatment_Jaw")
            $MandiblePreopElementv21.SetAttribute("JawId","Lower")
            $MandiblePreopElementv21.SetAttribute("Index",$USToothNumber)
            $MandiblePreopElementv21.SetAttribute("FileName","Mandible_PreOP.ply")
            $MandiblePreopElementv22 = $xmlv22.iTeroExport.ExportedObjects.AppendChild($xmlv22.CreateElement("Object"))
            $MandiblePreopElementv22.SetAttribute("ObjectType","Surface")
            $MandiblePreopElementv22.SetAttribute("SubType","PreTreatment_Jaw")
            $MandiblePreopElementv22.SetAttribute("JawId","Lower")
            $MandiblePreopElementv22.SetAttribute("Index",$USToothNumber)
            $MandiblePreopElementv22.SetAttribute("FileName","Mandible_PreOP.ply")
            $MandiblePreopElementv23 = $xmlv23.iTeroExport.ExportedObjects.AppendChild($xmlv23.CreateElement("Object"))
            $MandiblePreopElementv23.SetAttribute("ObjectType","Surface")
            $MandiblePreopElementv23.SetAttribute("SubType","PreTreatment_Jaw")
            $MandiblePreopElementv23.SetAttribute("JawId","Lower")
            $MandiblePreopElementv23.SetAttribute("Index",$USToothNumber)
            $MandiblePreopElementv23.SetAttribute("FileName","Mandible_PreOP.ply")
            $MandiblePreopElementv24 = $xmlv24.iTeroExport.ExportedObjects.AppendChild($xmlv24.CreateElement("Object"))
            $MandiblePreopElementv24.SetAttribute("ObjectType","Surface")
            $MandiblePreopElementv24.SetAttribute("SubType","PreTreatment_Jaw")
            $MandiblePreopElementv24.SetAttribute("JawId","Lower")
            $MandiblePreopElementv24.SetAttribute("Index",$USToothNumber)
            $MandiblePreopElementv24.SetAttribute("FileName","Mandible_PreOP.ply")
            $MandiblePreopElementv50 = $xmlv50.iTeroExport.ExportedObjects.AppendChild($xmlv50.CreateElement("Object"))
            $MandiblePreopElementv50.SetAttribute("ObjectType","Surface")
            $MandiblePreopElementv50.SetAttribute("SubType","PreTreatment_Jaw")
            $MandiblePreopElementv50.SetAttribute("JawId","Lower")
            $MandiblePreopElementv50.SetAttribute("Index",$USToothNumber)
            $MandiblePreopElementv50.SetAttribute("FileName","Mandible_PreOP.ply")
        }

        ###Update XML files: Add marin file info.
        ###Since $MarginFiles can be string if margin file does not exist (earlier 'if' statement), only add margin element based of $MarginFiles list type.
        if($MarginFiles.GetType().Name -eq 'String'){
            }else{
            $NewMarginElementv10 = $xmlv10.iTeroExport.ExportedObjects.AppendChild($xmlv10.CreateElement("Object"))
            $NewMarginElementv10.SetAttribute("ObjectType","Line")
            $NewMarginElementv10.SetAttribute("SubType","Margin")
            $NewMarginElementv10.SetAttribute("AdaId",$USToothNumber)
            $NewMarginElementv10.SetAttribute("FileName",$MarginFile.Name)
            $NewMarginElementv21 = $xmlv21.iTeroExport.ExportedObjects.AppendChild($xmlv21.CreateElement("Object"))
            $NewMarginElementv21.SetAttribute("ObjectType","Line")
            $NewMarginElementv21.SetAttribute("SubType","Margin")
            $NewMarginElementv21.SetAttribute("AdaId",$USToothNumber)
            $NewMarginElementv21.SetAttribute("FileName",$MarginFile.Name)
            $NewMarginElementv22 = $xmlv22.iTeroExport.ExportedObjects.AppendChild($xmlv22.CreateElement("Object"))
            $NewMarginElementv22.SetAttribute("ObjectType","Line")
            $NewMarginElementv22.SetAttribute("SubType","Margin")
            $NewMarginElementv22.SetAttribute("AdaId",$USToothNumber)
            $NewMarginElementv22.SetAttribute("FileName",$MarginFile.Name)
            $NewMarginElementv23 = $xmlv23.iTeroExport.ExportedObjects.AppendChild($xmlv23.CreateElement("Object"))
            $NewMarginElementv23.SetAttribute("ObjectType","Line")
            $NewMarginElementv23.SetAttribute("SubType","Margin")
            $NewMarginElementv23.SetAttribute("AdaId",$USToothNumber)
            $NewMarginElementv23.SetAttribute("FileName",$MarginFile.Name)
            $NewMarginElementv24 = $xmlv24.iTeroExport.ExportedObjects.AppendChild($xmlv24.CreateElement("Object"))
            $NewMarginElementv24.SetAttribute("ObjectType","Line")
            $NewMarginElementv24.SetAttribute("SubType","Margin")
            $NewMarginElementv24.SetAttribute("AdaId",$USToothNumber)
            $NewMarginElementv24.SetAttribute("FileName",$MarginFile.Name)
            $NewMarginElementv50 = $xmlv50.iTeroExport.ExportedObjects.AppendChild($xmlv50.CreateElement("Object"))
            $NewMarginElementv50.SetAttribute("ObjectType","Line")
            $NewMarginElementv50.SetAttribute("SubType","Margin")
            $NewMarginElementv50.SetAttribute("AdaId",$USToothNumber)
            $NewMarginElementv50.SetAttribute("FileName",$MarginFile.Name)
            }
        }

        ###Update XML with upper/lower jaw file names.
        ###Since preop element also contains 'JawId', include 'SubType' to isolate correct element to update)
        $nodev10 = $xmlv10.iTeroExport.ExportedObjects.Object |
        where {($_.JawId -eq 'Lower') -and ($_.SubType -eq 'Jaw')}
        $nodev10.FileName = $LowerJawScan
        $nodev21 = $xmlv21.iTeroExport.ExportedObjects.Object |
        where {($_.JawId -eq 'Lower') -and ($_.SubType -eq 'Jaw')}
        $nodev21.FileName = $LowerJawScan
        $nodev22 = $xmlv22.iTeroExport.ExportedObjects.Object |
        where {($_.JawId -eq 'Lower') -and ($_.SubType -eq 'Jaw')}
        $nodev22.FileName = $LowerJawScan
        $nodev23 = $xmlv23.iTeroExport.ExportedObjects.Object |
        where {($_.JawId -eq 'Lower') -and ($_.SubType -eq 'Jaw')}
        $nodev23.FileName = $LowerJawScan
        $nodev24 = $xmlv24.iTeroExport.ExportedObjects.Object |
        where {($_.JawId -eq 'Lower') -and ($_.SubType -eq 'Jaw')}
        $nodev24.FileName = $LowerJawScan
        $nodev50 = $xmlv50.iTeroExport.ExportedObjects.Object |
        where {($_.JawId -eq 'Lower') -and ($_.SubType -eq 'Jaw')}
        $nodev50.FileName = $LowerJawScan

        $nodev10 = $xmlv10.iTeroExport.ExportedObjects.Object |
        where {($_.JawId -eq 'Upper') -and ($_.SubType -eq 'Jaw')}
        $nodev10.FileName = $UpperJawScan
        $nodev21 = $xmlv21.iTeroExport.ExportedObjects.Object |
        where {($_.JawId -eq 'Upper') -and ($_.SubType -eq 'Jaw')}
        $nodev21.FileName = $UpperJawScan
        $nodev22 = $xmlv22.iTeroExport.ExportedObjects.Object |
        where {($_.JawId -eq 'Upper') -and ($_.SubType -eq 'Jaw')}
        $nodev22.FileName = $UpperJawScan
        $nodev23 = $xmlv23.iTeroExport.ExportedObjects.Object |
        where {($_.JawId -eq 'Upper') -and ($_.SubType -eq 'Jaw')}
        $nodev23.FileName = $UpperJawScan
        $nodev24 = $xmlv24.iTeroExport.ExportedObjects.Object |
        where {($_.JawId -eq 'Upper') -and ($_.SubType -eq 'Jaw')}
        $nodev24.FileName = $UpperJawScan
        $nodev50 = $xmlv50.iTeroExport.ExportedObjects.Object |
        where {($_.JawId -eq 'Upper') -and ($_.SubType -eq 'Jaw')}
        $nodev50.FileName = $UpperJawScan

        ###Save each XML document.
        $xmlv10.Save($pathv10)
        $xmlv21.Save($pathv21)
        $xmlv22.Save($pathv22)
        $xmlv23.Save($pathv23)
        $xmlv24.Save($pathv24)
        $xmlv50.Save($pathv50)
}