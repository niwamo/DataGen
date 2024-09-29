$classes = Get-ChildItem -Path "$PSScriptRoot/Types" -File -Filter "*.ps1"
foreach ($class in $classes) {
    . $class.FullName
}

function New-DummyData {
    param (
        [int]$NumFiles = 1,
        [int]$WordsPerFile = 1000,
        [string]$OutputFolder = "./out",
        [int]$NumFolders = 1,
        [int]$ThreadsPerApp = 1,
        [hashtable]$FileTypeRatio = @{
            txt=.25;
            docx=.25;
            pptx=.25;
            xlsx=.25
        },
        [ValidateScript({Test-Path $_})]
        [string]$Dict = "$PSScriptRoot/Data/dict.json",
        [ValidateScript({ (0 -le $_) -and ($_ -le 100) })]
        [int]$PercentSensitive = 0,
        [ValidateScript({Test-Path $_})]
        [string]$SensitiveData = "$PSScriptRoot/Data/sensitive-data.csv"
    )
    # read in data 
    $dictdata = Get-Content $Dict | ConvertFrom-Json
    $d = [Collections.Generic.Dictionary[string, int]]::new()
    foreach ($prop in ($dictdata | Get-Member -MemberType NoteProperty).Name) {
        $d[$prop] = $data.$prop
    }
    $DICTIONARY = [FrequencyDict]::new($d)
    $s = Get-Content -Path $SensitiveData | ConvertFrom-Csv 
    $SENSITIVE = [SText]::new($s)

    # make folders
    $Folders = New-FolderSet -OutputFolder $OutputFolder -NumFolders $NumFolders -Dictionary $DICTIONARY
    # get filetype counts
    $FileTypeCounts = @{}
    foreach ($type in "docx","pptx","xlsx") {
        $TotalCount = [int][Math]::Floor($FileTypeRatio[$type] * $NumFiles)
        if ($TotalCount -gt 0) {
            $FileTypeCounts.Add($type, (Get-NumSplit -Total $TotalCount -SplitCount $ThreadsPerApp))
        }
    }
    $Remainder = $NumFiles - `
        ($FileTypeCounts.Values | ForEach-Object { ($_ | Measure-Object -Sum).Sum } | Measure-Object -Sum).Sum
    $FileTypeCounts.Add(
        "txt", 
        (Get-NumSplit -Total $Remainder -SplitCount $ThreadsPerApp)
    )
    # Populate folders
    $Seed = Get-Random
    foreach ($type in $FileTypeCounts.Keys) {
        $preExisting = Get-Process | Where-Object Name -match "excel|power|word"
        foreach ($filecount in $FileTypeCounts[$type]) {
            $Seed += 1
            Start-Job -ArgumentList `
                    $type, `
                    $filecount, `
                    $WordsPerFile, `
                    $Folders, `
                    $DICTIONARY.ToString(), `
                    $PSCommandPath, `
                    $Seed, `
                    $PercentSensitive, `
                    $SENSITIVE.ToString() `
                -ScriptBlock {
                    param (
                        [string]$type,
                        [int]$filecount,
                        [int]$WordsPerFile,
                        [string[]]$Folders,
                        [string]$DictString,
                        [string]$ModulePath,
                        [int]$Seed,
                        [int]$PercentSensitive,
                        [string]$SensString
                    )
                    Import-Module $ModulePath
                    $classes = Get-ChildItem -Path "$ModulePath/../Types" -File -Filter "*.ps1"
                    foreach ($class in $classes) {
                        . $class.FullName
                    }
                    $Dictionary = New-FDFromString -Data $DictString
                    $Sensitive = New-STFromString -Data $SensString
                    if ($type -eq "docx") {
                        New-DocxFiles -FileCount $filecount -WordCount $WordsPerFile `
                            -Folders $Folders -Dictionary $DICTIONARY -Seed $Seed `
                            -PercentSensitive $PercentSensitive -Sensitive $Sensitive
                    } elseif ($type -eq "pptx") {
                        New-PptxFiles -FileCount $filecount -WordCount $WordsPerFile `
                            -Folders $Folders -Dictionary $DICTIONARY -Seed $Seed `
                            -PercentSensitive $PercentSensitive -Sensitive $Sensitive
                    } elseif ($type -eq "xlsx") {
                        New-XlsxFiles -FileCount $filecount -WordCount $WordsPerFile `
                            -Folders $Folders -Dictionary $DICTIONARY -Seed $Seed `
                            -PercentSensitive $PercentSensitive -Sensitive $Sensitive
                    } elseif ($type -eq "txt") {
                        New-TxtFiles -FileCount $filecount -WordCount $WordsPerFile `
                            -Folders $Folders -Dictionary $DICTIONARY -Seed $Seed `
                            -PercentSensitive $PercentSensitive -Sensitive $Sensitive
                    } else {
                        Throw "Unrecognized filetype. No file creator defined."
                    }
                }
        }
    }
    foreach ($job in Get-Job) {
        $job | Wait-Job -Force
        $job | Receive-Job
    }
    [GC]::Collect()
    foreach ($p in (Get-Process | Where-Object Name -match "excel|power|word")) {
        if ($p.Id -notin $preExisting.Id) {
            Stop-Process $p.Id
        }
    }
}

function New-FolderSet {
    param (
        [string]$OutputFolder,
        [int]$NumFolders,
        [FrequencyDict]$Dictionary
    )
    # ensure output folder exists
    if (-not (Test-Path -Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory
    }
    # make sure we have the full path
    $OutputFolder = Resolve-Path $OutputFolder | Select-Object -ExpandProperty Path
    $folders = New-Object string[] $NumFolders
    $folders[0] = $OutputFolder
    $generator = New-Selector
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
        [int]$FileCount,
        [int]$WordCount,
        [string[]]$Folders,
        [FrequencyDict]$Dictionary,
        [int]$Seed,
        [int]$PercentSensitive,
        [SText]$Sensitive
    )
    $POWERPOINT = New-Object -ComObject PowerPoint.Application
    $Generator = New-Selector -Seed $Seed
    $Extension = ".pptx"
    $Words = $Generator.select($Dictionary, $FileCount*2)
    $Dirs = $Generator.select($Folders, $FileCount)
    $OutputPaths = for ($i=0; $i -lt $FileCount; $i++) {
        $start = 2*$i
        $name = ($Words[$start..($start+1)] -join '-') + $Extension
        Join-Path -Path $Dirs[$i] -ChildPath $name
    }
    $wordsPerSlide = 200
    foreach ($OutPath in $OutputPaths) {
        # Create a new presentation
        try {
            $presentation = $POWERPOINT.Presentations.Add($false)
        } catch {
            $_.ErrorDetails = "Could not creat presentation"
            throw $_
        }

        # Add slides to the presentation
        $words = $Generator.select($Dictionary, $WordCount)
        $sOut = $false
        if ($generator.num(1, 101) -le $PercentSensitive) {
            # inject sensitive data
            $sOut = $true
            $mid = $generator.num(1, $WordCount)
            $words = `
                $words[0..$mid] + `
                $Sensitive.GetText($generator.generator) + `
                $words[($mid+1)..($WordCount-1)]
        }

        $SlideCount = [Math]::Max([int][Math]::Floor($words.Count / $wordsPerSlide), 1)
        for ($i = 0; $i -lt $SlideCount; $i++) {
            $slide = $presentation.Slides.Add($i + 1, 1)
            # msoTextOrientationHorizontal = 1
            $content = $slide.Shapes.AddTextbox(
                1, 10, 10, 900, 300
            )
            $start = $i * $wordsPerSlide
            $end = [Math]::Min($start + ($wordsPerSlide - 1), $words.Count)
            $content.TextFrame.TextRange.Text = $Words[$start..$end] -join ' '
        }
        :jail for ($attempts=0; $attempts -lt 10; $attempts++) {
            try {
                # Save the presentation
                $presentation.SaveAs([string]$OutPath)
                # Close PowerPoint
                $presentation.Close()
                if ($sOut) {
                    Write-Host "Injected sensitive data into $OutPath"
                }
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
    [GC]::Collect()
}

function New-DocxFiles {
    param (
        [int]$FileCount,
        [int]$WordCount,
        [string[]]$Folders,
        [FrequencyDict]$Dictionary,
        [int]$Seed, 
        [int]$PercentSensitive,
        [SText]$Sensitive
    )
    $WORD = New-Object -ComObject Word.Application
    $Generator = New-Selector -Seed $Seed
    
    $Extension = ".docx"
    $Words = $Generator.select($Dictionary, $FileCount*2)
    $Dirs = $Generator.select($Folders, $FileCount)
    $OutputPaths = for ($i=0; $i -lt $FileCount; $i++) {
        $start = 2*$i
        $name = ($Words[$start..($start+1)] -join '-') + $Extension
        Join-Path -Path $Dirs[$i] -ChildPath $name
    }
    
    $WordCount += 1
    foreach ($OutPath in $OutputPaths) {
        # Create a new file
        try {
            $document = $WORD.Documents.Add()
        } catch {
            $_.ErrorDetails = "Could not create the document"
            throw $_
        }
        
        # get content
        $words = $Generator.select($Dictionary, $WordCount)
        $sOut = $false
        if ($generator.num(1, 101) -le $PercentSensitive) {
            # inject sensitive data
            $sOut = $true
            $mid = $generator.num(1, $WordCount)
            $words = `
                $words[0..$mid] + `
                $Sensitive.GetText($generator.generator) + `
                $words[($mid+1)..($WordCount-1)]
        }

        # Add a title to the document
        $p = $document.Content.Paragraphs.Add()
        $p.Range.Text = $words[0]

        # Add content to the document
        $p = $document.Content.Paragraphs.Add()
        $p.Range.Text = ($words[1..$words.Count] -join ' ')

        :jail for ($attempts=0; $attempts -lt 10; $attempts++) {
            try {
                # Save the file
                $document.SaveAs([string]$OutPath)
                # Close the file
                $document.Close()
                if ($sOut) {
                    Write-Host "Injected sensitive data into $OutPath"
                }
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
    [GC]::Collect()
}

function New-XlsxFiles {
    param (
        [int]$FileCount,
        [int]$WordCount,
        [string[]]$Folders,
        [FrequencyDict]$Dictionary,
        [int]$Seed,
        [int]$PercentSensitive,
        [SText]$Sensitive
    )
    $EXCEL = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Generator = New-Selector -Seed $Seed
    
    $Extension = ".xlsx"
    $Words = $Generator.select($Dictionary, $FileCount*2)
    $Dirs = $Generator.select($Folders, $FileCount)
    $OutputPaths = for ($i=0; $i -lt $FileCount; $i++) {
        $start = 2*$i
        $name = ($Words[$start..($start+1)] -join '-') + $Extension
        Join-Path -Path $Dirs[$i] -ChildPath $name
    }
    
    $numCols = 50
    foreach ($OutPath in $OutputPaths) {
        $WordCount += 1
        $words = $Generator.select($Dictionary, $WordCount)
        $start = 0; $end = $WordCount-1
        $sOut = $false
        if ($generator.num(1, 101) -le $PercentSensitive) {
            # inject sensitive data
            $sOut = $true
            $mid = $generator.num(1, $WordCount)
            $words = `
                $words[0..$mid] + `
                $Sensitive.GetText($generator.generator) + `
                $words[($mid+1)..($WordCount-1)]
        }

        # Add a new workbook
        try {
            $workbook = $EXCEL.Workbooks.Add()
        } catch {
            $_.ErrorDetails = "Could not create workbook"
            throw $_
        }

        # Select the first sheet
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = $words[0]

        $RowCount = [Math]::Max([int][Math]::Ceiling($words.Count / $numCols), 1)

        for ($row=1; $row -le $RowCount; $row++) {
            $startCell = $sheet.Cells.Item($row, 1)
            $endCell = $sheet.Cells.Item(
                $row,
                [Math]::Min($numCols, $words.Count)
            )

            # Assign the data to the range
            $range = $sheet.Range($startCell, $endCell)
            $start = ($row - 1) * $numCols
            $end = [Math]::Min($start + ($numCols-1), $words.Count)
            $range.Value = @($words[$start..$end])
        }
                    
        :jail for ($attempts=0; $attempts -lt 10; $attempts++) {
            try {
                # Save the file
                $workbook.SaveCopyAs([string]$OutPath)
                $workbook.Close($false)
                if ($sOut) {
                    Write-Host "Injected sensitive data into $OutPath"
                }
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

function New-TxtFiles {
    param (
        [int]$FileCount,
        [int]$WordCount,
        [string[]]$Folders,
        [FrequencyDict]$Dictionary,
        [int]$Seed,
        [int]$PercentSensitive,
        [SText]$Sensitive

    )
    $generator = New-Selector -Seed $Seed
    $extension = ".txt"
    $words = $generator.select($Dictionary, $FileCount*2)
    $dirs = $generator.select($Folders, $FileCount)
    $outputPaths = for ($i=0; $i -lt $FileCount; $i++) {
        $start = 2*$i
        $name = ($words[$start..($start+1)] -join '-') + $extension
        Join-Path -Path $dirs[$i] -ChildPath $name
    }
    foreach ($outPath in $outputPaths) {
        $words = $generator.select($Dictionary, $WordCount)
        $sOut = $false
        if ($generator.num(1, 101) -le $PercentSensitive) {
            # inject sensitive data
            $sOut = $true
            $mid = $generator.num(1, $WordCount-1)
            $words = `
                $words[0..$mid] + `
                $Sensitive.GetText($generator.generator) + `
                $words[($mid+1)..($WordCount-1)]
        }
        :jail for ($attempts=0; $attempts -lt 10; $attempts++) {
            try {
                Set-Content -Path $OutPath -Value ($words -join ' ')
                if ($sOut) {
                    Write-Host "Injected sensitive data into $OutPath"
                }
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
