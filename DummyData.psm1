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
    $Remainder = $NumFiles - `
        ($FileTypeCounts.Values | ForEach-Object { ($_ | Measure-Object -Sum).Sum } | Measure-Object -Sum).Sum
    $FileTypeCounts.Add(
        "txt", 
        (Get-NumSplit -Total $Remainder -SplitCount $ThreadsPerApp)
    )
    
    # Populate folders
    $Seed = 1
    foreach ($type in $FileTypeCounts.Keys) {
        $jobs = foreach ($filecount in $FileTypeCounts[$type]) {
            Start-Job -ArgumentList $type, $filecount, $WordsPerFile, $Folders, $Dictionary, $PSCommandPath, $Seed `
                -ScriptBlock {
                    param (
                        [string]$type,
                        [int]$filecount,
                        [int]$WordsPerFile,
                        [string[]]$Folders,
                        [string[]]$Dictionary,
                        [string]$ModulePath,
                        [int]$Seed
                    )
                    Import-Module $ModulePath
                    if ($type -eq "docx") {
                        New-DocxFiles -Count $filecount -WordCount $WordsPerFile -Folders $Folders -Dictionary $DICTIONARY -Seed $Seed
                    } elseif ($type -eq "pptx") {
                        New-PptxFiles -Count $filecount -WordCount $WordsPerFile -Folders $Folders -Dictionary $DICTIONARY -Seed $Seed
                    } elseif ($type -eq "xlsx") {
                        New-XlsxFiles -Count $filecount -WordCount $WordsPerFile -Folders $Folders -Dictionary $DICTIONARY -Seed $Seed
                    } elseif ($type -eq "txt") {
                        New-TxtFiles -Count $filecount -WordCount $WordsPerFile -Folders $Folders -Dictionary $DICTIONARY -Seed $Seed
                    } else {
                        Throw "Unrecognized filetype. No file creator defined."
                    }
                }
            $Seed++
        }
    }
    foreach ($job in $jobs) {
        $job | Wait-Job
        $job | Receive-Job
    }
    [GC]::Collect()
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
    Randomizer ([int]$Seed) {
        $this.generator = New-Object System.Random($Seed)
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

function New-PptxFiles {
    param (
        [int]$Count,
        [int]$WordCount,
        [string[]]$Folders,
        [string[]]$Dictionary,
        [int]$Seed
    )
        
    & { 
        $POWERPOINT = New-Object -ComObject PowerPoint.Application
        $Generator = New-Object Randomizer($Seed)
        
        $Extension = ".pptx"
        $Words = $Generator.select($Dictionary, $Count*2)
        $Dirs = $Generator.select($Folders, $Count)
        $OutputPaths = for ($i=0; $i -lt $Count; $i++) {
            $start = 2*$i
            $name = ($Words[$start..($start+1)] -join '-') + $Extension
            Join-Path -Path $Dirs[$i] -ChildPath $name
        }
        
        foreach ($OutPath in $OutputPaths) {
            # Create a new presentation
            $presentation = $POWERPOINT.Presentations.Add()

            # Add slides to the presentation
            $Words = $Generator.select($Dictionary, $WordCount)
            $SlideCount = [Math]::Max([int][Math]::Floor($WordCount / 500), 1)
            for ($i = 0; $i -lt $SlideCount; $i++) {
                $slide = $presentation.Slides.Add($i + 1, 1)
                $content = $slide.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 50, 100, 600, 300)
                $start = $i * 500
                $end = [Math]::Min($start + 499, $WordCount)
                $content.TextFrame.TextRange.Text = $Words[$start..$end] -join ' '
            }
            :jail for ($attempts=0; $attempts -lt 10; $attempts++) {
                try {
                    # Save the presentation
                    $presentation.SaveAs([string]$OutPath)

                    # Close PowerPoint
                    $presentation.Close()

                    # Write-Output "PowerPoint presentation created: $OutputPath"
                    break jail
                }
                catch {
                    Write-Error "Error creating PowerPoint presentation: $_"
                    $name = ($Generator.select($Dictionary, 2) -join '-') + $Extension
                    $OutPath = Join-Path -Path $Generator.select($Folders) -ChildPath $name
                }
            }
        }

        $POWERPOINT.Quit()
    }
    [GC]::Collect()
}

function New-DocxFiles {
    param (
        [int]$Count,
        [int]$WordCount,
        [string[]]$Folders,
        [string[]]$Dictionary,
        [int]$Seed
    )
        
    & { 
        $WORD = New-Object -ComObject Word.Application
        $Generator = New-Object Randomizer($Seed)
        
        $Extension = ".docx"
        $Words = $Generator.select($Dictionary, $Count*2)
        $Dirs = $Generator.select($Folders, $Count)
        $OutputPaths = for ($i=0; $i -lt $Count; $i++) {
            $start = 2*$i
            $name = ($Words[$start..($start+1)] -join '-') + $Extension
            Join-Path -Path $Dirs[$i] -ChildPath $name
        }
        
        foreach ($OutPath in $OutputPaths) {
            # Create a new file
            $document = $WORD.Documents.Add()

            $words = $Generator.select($Dictionary, $WordCount+1)

            # Add a title to the document
            $document.Content.Text = $words[0]

            # Add content to the document
            $document.Content.InsertParagraphAfter()
            $document.Content.Text = $words[1..$WordCount] -join ' '

            :jail for ($attempts=0; $attempts -lt 10; $attempts++) {
                try {
                    # Save the file
                    $document.SaveAs([string]$OutPath)
                    # Close the file
                    $document.Close()
                    break jail
                }
                catch {
                    Write-Error "Error creating file: $_"
                    $name = ($Generator.select($Dictionary, 2) -join '-') + $Extension
                    $OutPath = Join-Path -Path $Generator.select($Folders) -ChildPath $name
                }
            }
        }
        $WORD.Quit()
    }
    [GC]::Collect()
}

function New-XlsxFiles {
    param (
        [int]$Count,
        [int]$WordCount,
        [string[]]$Folders,
        [string[]]$Dictionary,
        [int]$Seed
    )
        
    & { 
        $EXCEL = New-Object -ComObject Excel.Application
        $Generator = New-Object Randomizer($Seed)
        
        $Extension = ".xlsx"
        $Words = $Generator.select($Dictionary, $Count*2)
        $Dirs = $Generator.select($Folders, $Count)
        $OutputPaths = for ($i=0; $i -lt $Count; $i++) {
            $start = 2*$i
            $name = ($Words[$start..($start+1)] -join '-') + $Extension
            Join-Path -Path $Dirs[$i] -ChildPath $name
        }
        
        foreach ($OutPath in $OutputPaths) {
            $words = $Generator.select($Dictionary, $WordCount+1)

            # Add a new workbook
            $workbook = $EXCEL.Workbooks.Add()

            # Select the first sheet
            $sheet = $workbook.Worksheets.Item(1)
            $sheet.Name = $words[0]

            $RowCount = [Math]::Max([int][Math]::Floor($WordCount / 500), 1)

            for ($row=1; $row -le $RowCount; $row++) {
                $startCell = $sheet.Cells.Item($row, 1)
                $endCell = $sheet.Cells.Item(
                    $row,
                    [Math]::Min(500, $WordCount)
                )

                # Assign the data to the range
                $range = $sheet.Range($startCell, $endCell)
                $start = ($row - 1) * 500
                $end = [Math]::Min($start + 499, $WordCount)
                $range.Value = @($words[$start..$end])
            }
                     
            :jail for ($attempts=0; $attempts -lt 10; $attempts++) {
                try {
                    # Save the file
                    $workbook.SaveAs([string]$OutPath)
                    # Close the file
                    $workbook.Close()
                    break jail
                }
                catch {
                    Write-Error "Error creating file: $_"
                    $name = ($Generator.select($Dictionary, 2) -join '-') + $Extension
                    $OutPath = Join-Path -Path $Generator.select($Folders) -ChildPath $name
                }
            }
        }
        $EXCEL.Quit()
    }
    [GC]::Collect()
}

function New-TxtFiles {
    param (
        [int]$Count,
        [int]$WordCount,
        [string[]]$Folders,
        [string[]]$Dictionary,
        [int]$Seed
    )
        
    $Generator = New-Object Randomizer($Seed)
    
    $Extension = ".txt"
    $Words = $Generator.select($Dictionary, $Count*2)
    $Dirs = $Generator.select($Folders, $Count)
    $OutputPaths = for ($i=0; $i -lt $Count; $i++) {
        $start = 2*$i
        $name = ($Words[$start..($start+1)] -join '-') + $Extension
        Join-Path -Path $Dirs[$i] -ChildPath $name
    }
    
    foreach ($OutPath in $OutputPaths) {
        $words = $Generator.select($Dictionary, $WordCount+1)   
        :jail for ($attempts=0; $attempts -lt 10; $attempts++) {
            try {
                # Save the file
                Set-Content -Path $OutPath -Value ($words -join ' ')
                break jail
            }
            catch {
                Write-Error "Error creating file: $_"
                $name = ($Generator.select($Dictionary, 2) -join '-') + $Extension
                $OutPath = Join-Path -Path $Generator.select($Folders) -ChildPath $name
            }
        }
    }
}
