function New-DummyData {
    param (
        [int]$TotalFiles,
        [int]$WordsPerFile,
        [string]$OutputFolder,
        [int]$NestingLevel,
        [int]$FoldersPerLevel
    )
    
    # TODO: change folder generation from ridgid and structure to random 
        # create list of folders, randomly select one to 
    #

    <# Validate file type ratio
    $totalRatio = $FileTypeRatio.Values | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    if ($totalRatio -ne 100) {
        Throw "File type ratio must sum to 100"
    }
    #>

    # ensure output folder exists
    if (-not (Test-Path -Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory
    }
    $OutputFolder = Resolve-Path $OutputFolder | Select-Object -ExpandProperty Path

    <# Calculate the number of files for each type
    $fileCounts = @{}
    foreach ($fileType in $FileTypeRatio.Keys) {
        $count = [math]::Round($TotalFiles * ($FileTypeRatio[$fileType]/100))
        $fileCounts[$fileType] = $count
    }
    #>

    # Calculate the number of files per folder
    $TotalFolders = ([Math]::Pow($FoldersPerLevel,$NestingLevel+1) - 1)/($FoldersPerLevel - 1)
    $FilesPerFolder = [int][Math]::Floor($TotalFiles/$TotalFolders)
    $FilesPerTopFolder = $FilesPerFolder + ($TotalFiles % $TotalFolders)

    # Create the folder structure
    $Folders = Set-ChildFolders -ParentDir $OutputFolder -NumFolders $FoldersPerLevel `
        -MaxLevel $NestingLevel

    # Populate folders
    Set-FolderFiles -Folder $OutputFolder -NumFiles $FilesPerTopFolder -WordsPerFile $WordsPerFile
    foreach ($dir in $Folders) {
        Set-FolderFiles -Folder $dir -NumFiles $FilesPerFolder -WordsPerFile $WordsPerFile
    }
        
}

function Write-Green {
    process { Write-Host $_ -ForegroundColor Green }
}

function Get-MyRandom {
    param (
        [Object[]]$InputObject,
        [int]$Count = 1,
        [int]$Maximum
    )
    $random = New-Object System.Random
    if (! $PSBoundParameters.ContainsKey('InputObject')) {
        if (! $PSBoundParameters.ContainsKey('Maximum')) {
            Throw "Either InputObject or Maximum must be provided"
        }
        $out = for ($i=0; $i -lt $Count; $i++) {
            $random.Next(0, $Maximum)
        }
    } else {
        $n = $InputObject.Length - 1
        $out = for ($i=0; $i -lt $Count; $i++) { 
            $InputObject[($random.Next(0, $n))]
        }
    }
    return $out
}

function Set-ChildFolders {
    param (
        [string]$ParentDir,
        [int]$NumFolders,
        [int]$CurrentLevel,
        [int]$MaxLevel
    )

    $Folders = @()
    for ($i=0; $i -lt $NumFolders; $i++) {
        $FolderName = (Get-RandomWords -NumWords 2) -join '_'
        $FolderPath = Join-Path -Path $ParentDir -ChildPath $FolderName
        New-Item -ItemType Directory -Path $FolderPath
        $Folders += $FolderPath
    }

    if ( ! $PSBoundParameters.ContainsKey('CurrentLevel') ) {
        $CurrentLevel = 1
    }

    if ( $CurrentLevel -lt $MaxLevel ) {
        foreach ($ChildFolder in $Folders) {
            $Descendants = Set-ChildFolders -ParentDir $ChildFolder -NumFolders $NumFolders `
                -CurrentLevel ($CurrentLevel + 1) -MaxLevel $MaxLevel
            $Folders += $Descendants
        }
    }

    return $Folders
}

function Set-FolderFiles {
    param (
        [string]$Folder,
        [int]$NumFiles,
        [int]$WordsPerFile
    )

    # Write-Output "Adding $NumFiles to $Folder"

    for ($i = 0; $i -lt $NumFiles; $i++) {
        $extension = Get-Random -InputObject $FILE_TYPES
        $fileName = ((Get-RandomWords -NumWords 3) -join '-') + $extension
        $filePath = Join-Path -Path $Folder -ChildPath $fileName

        New-Document -Path $filePath -Type $extension -NumWords $WordsPerFile
        <# Optionally, zip files based on the ZipProportion
        if ((Get-Random -Minimum 0 -Maximum 100) -lt $ZipProportion) {
            Compress-Archive -Path $filePath -DestinationPath "$filePath.zip" -Force
            Remove-Item -Path $filePath -Force
        }
        #>
    }
}

function New-Document {
    param (
        [string]$Path,
        [string]$Type,
        [int]$NumWords
    )

    if ($Type -eq ".pptx") {
        New-PowerPointPresentation -OutputPath $Path -Title (Get-RandomWords -NumWords 1) `
            -SlideTitles (Get-RandomWords -NumWords 5) `
            -SlideContents @(
                (Get-RandomWords -NumWords ($NumWords/5)), 
                (Get-RandomWords -NumWords ($NumWords/5)), 
                (Get-RandomWords -NumWords ($NumWords/5)), 
                (Get-RandomWords -NumWords ($NumWords/5)), 
                (Get-RandomWords -NumWords ($NumWords/5)) 
            )
    } elseif ($Type -eq ".docx") {
        New-WordDocument -OutputPath $Path -Title (Get-RandomWords -NumWords 1) `
            -Content ((Get-RandomWords -NumWords $NumWords) -join ' ')
    } elseif ($Type -eq ".xlsx") {
        New-ExcelDocument -OutputPath $Path -SheetName (Get-RandomWords -NumWords 1) `
            -Headers (Get-RandomWords -NumWords 5) `
            -Data @(
                (Get-RandomWords -NumWords ($NumWords/5)), 
                (Get-RandomWords -NumWords ($NumWords/5)), 
                (Get-RandomWords -NumWords ($NumWords/5)), 
                (Get-RandomWords -NumWords ($NumWords/5)), 
                (Get-RandomWords -NumWords ($NumWords/5)) 
            )
    } else {
        Set-Content -Path $Path -Value ((Get-RandomWords -NumWords $NumWords) -join ' ')
    }
}

function New-PowerPointPresentation {
    param (
        [string]$OutputPath,
        [string]$Title,
        [string[]]$SlideTitles,
        [string[]]$SlideContents
    )

    try {
        # Create a new presentation
        $presentation = $POWERPOINT.Presentations.Add()

        # Add slides to the presentation
        for ($i = 0; $i -lt $SlideTitles.Count; $i++) {
            $slide = $presentation.Slides.Add($i + 1, 1)
            $content = $slide.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 50, 100, 600, 300)
            $content.TextFrame.TextRange.Text = $SlideContents[$i]
        }

        # Save the presentation
        $presentation.SaveAs($OutputPath)

        # Close PowerPoint
        $presentation.Close()

        # Write-Host "PowerPoint presentation created: $OutputPath"
    }
    catch {
        Write-Error "Error creating PowerPoint presentation: $_"
    }
}

function New-WordDocument {
    param (
        [string]$OutputPath,
        [string]$Title,
        [string]$Content
    )

    try {
        # Add a new document
        $document = $WORD.Documents.Add()

        # Add a title to the document
        $document.Content.Text = $Title

        # Add content to the document
        $document.Content.InsertParagraphAfter()
        $document.Content.Text = $Content

        # Save the Word document
        $document.SaveAs($OutputPath)

        # Close Word
        $document.Close()

        # Write-Host "Word document created: $OutputPath"
    }
    catch {
        Write-Error "Error creating Word document: $_"
    }
}

function New-ExcelDocument {
    param (
        [string]$OutputPath,
        [string]$SheetName,
        [string[]]$Headers,
        [object[][]]$Data
    )
    try {
        
        # Add a new workbook
        $workbook = $EXCEL.Workbooks.Add()

        # Select the first sheet
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = $SheetName

        # Add headers to the sheet
        $row = 1
        $col = 1
        foreach ($header in $Headers) {
            $sheet.Cells.Item($row, $col) = $header
            $col++
        }

        # Add data to the sheet
        $row++
        $col = 1
        foreach ($rowData in $Data) {
            $row = 2
            foreach ($cellData in $rowData) {
                $sheet.Cells.Item($row, $col) = $cellData
                $row++
            }
            $col++
        }

        # Save the Excel workbook
        $workbook.SaveAs($OutputPath)

        # Close Excel
        $workbook.Close()

        # Write-Host "Excel document created: $OutputPath"
    }
    catch {
        Write-Error "Error creating Excel document: $_"
    }
}


<#
$DICT_URL = "https://github.com/dolph/dictionary/raw/master/popular.txt"
$DICTIONARY = (Invoke-WebRequest -URI $DICT_URL -UseBasicParsing).Content -split "`r`n"
#>
$DICTIONARY = Get-Content "./dictionary.txt"
$DICT_SIZE = $DICTIONARY | Measure-Object | Select-Object -ExpandProperty Count

$FILE_TYPES = @(".xlsx", ".pptx", ".docx", ".txt")

$POWERPOINT = New-Object -ComObject PowerPoint.Application
$WORD = New-Object -ComObject Word.Application
$EXCEL = New-Object -ComObject Excel.Application

New-DataRepo -TotalFiles 250000 -WordsPerFile 10000 -OutputFolder "./data-repo" -NestingLevel 5 -FoldersPerLevel 5

$EXCEL.Quit()
$WORD.Quit()
$POWERPOINT.Quit()