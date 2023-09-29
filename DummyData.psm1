function New-DummyData {
    param (
        [int]$NumFiles = 1,
        [int]$WordsPerFile = 1000,
        [string]$OutputFolder = "./out",
        [int]$NumFolders = 1,
        [int]$ThreadsPerApp = 1,
        [string[]]$DICTIONARY,
        [hashtable]$FileTypeRatio = @{
            txt=.25;
            docx=.25;
            pptx=.25;
            xlsx=.25
        }
    )
     
    # make sure we have a dictionary
    if (! $PSBoundParameters.ContainsKey('DICTIONARY')) {
        $DICT_URL = "https://github.com/dolph/dictionary/raw/master/popular.txt"
        $DICTIONARY = (Invoke-WebRequest -URI $DICT_URL -UseBasicParsing).Content -split "`n"
    }

    # make folders
    $Folders = New-FolderSet -OutputFolder $OutputFolder -NumFolders $NumFolders -Dictionary $DICTIONARY

    # get filetype counts

    



    $FileTypeCounts = @{}
    foreach ($type in "docx","pptx","xlsx") {
        $TotalCount = [int][Math]::Floor($FileTypeRatio[$type] * $NumFiles)
        $FileTypeCounts.Add($type, (Get-NumSplit -Total $TotalCount -SplitCount $ThreadsPerApp))
    }
    $Remainder = $NumFiles - ($FileTypeCounts.Values | Measure-Object -Sum).Sum
    $FileTypeCounts.Add(
        "txt", 
        (Get-NumSplit -Total $Remainder -SplitCount $ThreadsPerApp)
    )
    
    # TODO
    # Populate folders
    foreach ($type in $FileTypeCounts.Keys) {
        for ($i=0; $i -lt $ThreadsPerApp; $i++) {
            if ($type -eq "docx") {
                Start-ThreadJob -ScriptBlock {
                    New-DocxFiles -Count ($FileTypeCounts[$type])
                }
            } elseif ($type -eq "pptx") {

            } elseif ($type -eq "xlsx") {

            } elseif ($type -eq "txt") {

            } else {
                Throw "Unrecognized filetype. No file creator defined."
            }
            
        }
    }
}

function New-FolderSet {
    param (
        [string]$OutputFolder,
        [int]$NumFolders,
        [string[]]$Dictionary
    )
    # ensure output folder exists
    if (-not (Test-Path -Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory
    }
    # make sure we have the full path
    $OutputFolder = Resolve-Path $OutputFolder | Select-Object -ExpandProperty Path
    $folders = New-Object string[] $NumFolders
    $folders[0] = $OutputFolder
    $generator = New-Object Randomizer
    $names = $generator.select($Dictionary, $NumFolders)
    for ( $i=1; $i -lt $NumFolders; $i++ ) {
        $folders[$i] = Join-Path -Path $folders[$generator.num($i)] `
            -ChildPath $names[$i]
        try {
            [void](New-Item -ItemType Directory -Path $folders[$i])
        }
        catch {
            # if, by chance, we try to create something that already exists, retry
            $i--
        }
    }
    return $folders
}

class Randomizer {
    [System.Random]$generator
    Randomizer () {
        $this.generator = New-Object System.Random
    }
    [Object[]] select ([Object[]]$InputObject, [int]$Count) {
        $n = $InputObject.Length - 1
        $out = for ($i=0; $i -lt $Count; $i++) { 
            $InputObject[($this.generator.Next(0, $n))]
        }
        return $out
    }
    [Object] select ([Object[]]$InputObject) {
        $n = $InputObject.Length - 1
        return $InputObject[($this.generator.Next(0, $n))]
    }
    [int[]] num ([int]$Maximum, [int]$Count) {
        $out = for ($i=0; $i -lt $Count; $i++) {
            $this.generator.Next(0, $Maximum)
        }
        return $out
    }
    [int] num ([int]$Maximum) {
        return $this.generator.Next(0, $Maximum)
    }
}

function Get-NumSplit {
    param (
        [int]$Total,
        [int]$SplitCount
    )
    $counts = New-Object int[] $SplitCount
    for ($i=0; $i -lt ($SplitCount - 1); $i++) {
        $counts[$i] = [int][Math]::Round($Total/$SplitCount)
    }
    $counts[$SplitCount-1] = $Total - ($counts | Measure-Object -Sum).Sum
    return $counts
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

        # Write-Output "PowerPoint presentation created: $OutputPath"
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

        # Write-Output "Word document created: $OutputPath"
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

        # Write-Output "Excel document created: $OutputPath"
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

$FILE_TYPES = @(".xlsx", ".pptx", ".docx", ".txt")

$POWERPOINT = New-Object -ComObject PowerPoint.Application
$WORD = New-Object -ComObject Word.Application
$EXCEL = New-Object -ComObject Excel.Application

New-DataRepo -TotalFiles 250000 -WordsPerFile 10000 -OutputFolder "./data-repo" -NestingLevel 5 -FoldersPerLevel 5

$EXCEL.Quit()
$WORD.Quit()
$POWERPOINT.Quit()