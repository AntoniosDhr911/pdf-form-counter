Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Runtime.WindowsRuntime

# ============================================================
# SCHOOL FORM COUNTER
# SINGLE SCHOOL + BATCH MODE
#
# IMPORTANT:
# School folders are READ ONLY.
# This script never renames, moves, deletes, copies, edits,
# creates, or overwrites anything inside a school folder.
#
# Batch mode creates ONE report .txt on your Desktop.
#
# LOOSE / UNWORKED FORMS
#
# Forms outside Entered in D4w, Dups, E-forms and Paper forms
# are treated as loose/unworked forms.
#
# Examples:
#   School\Form 1.pdf
#   School\Late\Form 2.pdf
#   School\Forms from t1\Form 3.pdf
#
# LATE FOLDER RULE
#
# Folders named Late or Late forms are treated as unworked:
#   no UNSEPARATED or move-to-category warnings for their contents
#   PDF type is determined from pages, size and CDBS marker,
#   NEVER from the Late folder name
#   both e-forms and occasional paper forms are counted by type
#
# Loose forms:
#   ARE counted normally
#   are NOT flagged for naming
#   are NOT flagged as unseparated
#   are NOT flagged as being in the wrong E/Paper folder
#
# Once a form is placed into Entered in D4w, normal workflow
# and naming checks apply.
#
# DUPS RULE
#
# ANY recognized Dups folder anywhere under the school applies:
#
#   School\Dups
#   School\Forms from t1\Dups
#   School\Campus X\Dups
#   School\Campus Y\Dups
#
# E-forms inside ANY Dups folder are NOT counted.
#
# Paper forms inside ANY Dups folder:
#   ARE included in the displayed "Paper forms" number
#   are NOT included in "Total forms"
#
# Example:
#   Outside Dups: 30 E + 20 Paper = Total 50
#   Inside Dups:  10 E + 5 Paper
#
# Result:
#   Total forms = 50
#   E-forms     = 30
#   Paper forms = 25
#
# LOOSE COUNTS
#
# Loose forms outside Forms from t1 are reported as:
#   Loose forms
#
# Loose forms anywhere under Forms from t1 are reported as:
#   Forms from t1 loose
#
# If the school has no Forms from t1 folder, the report says so.
# ============================================================


# ---------- Windows PDF support ----------

$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]

$script:AsTaskMethod = (
    [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object {
        $_.Name -eq 'AsTask' -and
        $_.IsGenericMethod -and
        $_.GetParameters().Count -eq 1
    }
)[0]

function Wait-WinRT {
    param($Operation, [Type]$ResultType)

    $task = $script:AsTaskMethod.MakeGenericMethod($ResultType).Invoke(
        $null,
        @($Operation)
    )

    $task.Wait()
    return $task.Result
}

function Get-PdfPageCount {
    param([string]$Path)

    $file = Wait-WinRT `
        ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) `
        ([Windows.Storage.StorageFile])

    $pdf = Wait-WinRT `
        ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) `
        ([Windows.Data.Pdf.PdfDocument])

    return [int]$pdf.PageCount
}


# ---------- Folder-name recognition ----------

function Test-DupsFolderName {
    param([string]$Name)

    return (
        $Name -match
        '(?i)^(?:dup(?:s|e|es)?|duplicates?)(?:[\s_-]+forms?)?$'
    )
}

function Test-EFormsFolderName {
    param([string]$Name)
    return ($Name -match '(?i)^e[\s_-]*forms?$')
}

function Test-PaperFormsFolderName {
    param([string]$Name)
    return ($Name -match '(?i)^paper(?:[\s_-]*forms?)?$')
}

function Test-EnteredD4wFolderName {
    param([string]$Name)
    return ($Name -match '(?i)^entered[\s_-]*in[\s_-]*d4w$')
}

function Test-ConsentFormFolderName {
    param([string]$Name)
    return ($Name -match '(?i)consent')
}

function Test-LateFolderName {
    param([string]$Name)
    return ($Name -match '(?i)^late(?:[\s_-]*forms?)?$')
}

function Test-FormsFromT1FolderName {
    param([string]$Name)

    return (
        $Name -match '(?i)^forms?[\s_-]*from[\s_-]*t[\s_-]*1$'
    )
}


# ---------- Path classification ----------

function Get-LocationFlags {
    param([string]$RelativePath)

    $insideDups = $false
    $insideEntered = $false
    $insideEForms = $false
    $insidePaperForms = $false
    $insideFormsFromT1 = $false
    $insideLate = $false

    if (-not [string]::IsNullOrWhiteSpace($RelativePath)) {
        foreach ($part in ($RelativePath -split '\\')) {

            if (Test-DupsFolderName $part) {
                $insideDups = $true
            }

            if (Test-EnteredD4wFolderName $part) {
                $insideEntered = $true
            }

            if (Test-EFormsFolderName $part) {
                $insideEForms = $true
            }

            if (Test-PaperFormsFolderName $part) {
                $insidePaperForms = $true
            }

            if (Test-FormsFromT1FolderName $part) {
                $insideFormsFromT1 = $true
            }

            if (Test-LateFolderName $part) {
                $insideLate = $true
            }
        }
    }

    return [PSCustomObject]@{
        InsideDups        = $insideDups
        InsideEntered     = $insideEntered
        InsideEForms      = $insideEForms
        InsidePaperForms  = $insidePaperForms
        InsideFormsFromT1 = $insideFormsFromT1
        InsideLate        = $insideLate
    }
}


# ============================================================
# SCAN ONE SCHOOL
# ============================================================

function Scan-School {
    param([string]$SchoolPath)

    $schoolName = Split-Path $SchoolPath -Leaf

    $eForms = 0

    # Displayed Paper forms count:
    # includes paper forms everywhere, INCLUDING Dups.
    $paperForms = 0

    # Internal Total-forms paper count:
    # excludes paper forms inside ANY recognized Dups folder.
    $paperFormsForTotal = 0

    # Loose/unworked forms are still included in the normal totals.
    # These counters are informational only.
    $looseForms = 0
    $formsFromT1Loose = 0

    $mistakes = New-Object System.Collections.Generic.List[string]

    # NEW:
    # This list contains forms/folders left DIRECTLY inside
    # "Entered in D4w", next to E-forms and Paper forms.
    $unseparated = New-Object System.Collections.Generic.List[string]


    # ---------- Read folders ----------

    try {
        $allDirectories = @(
            Get-ChildItem `
                -LiteralPath $SchoolPath `
                -Directory `
                -Recurse `
                -ErrorAction Stop
        )
    }
    catch {
        $mistakes.Add(
            "COULD NOT READ SCHOOL FOLDERS: $($_.Exception.Message)"
        )
        $allDirectories = @()
    }


    # ---------- Detect Forms from t1 ----------

    $formsFromT1Folders = @(
        $allDirectories |
        Where-Object {
            Test-FormsFromT1FolderName $_.Name
        }
    )

    $hasFormsFromT1 = (
        $formsFromT1Folders.Count -gt 0
    )


    # ---------- Check standard folder names ----------

    foreach ($directory in $allDirectories) {

        $name = $directory.Name

        $relativeFolder = $directory.FullName.Substring(
            $SchoolPath.Length
        ).TrimStart('\')

        if (
            (Test-DupsFolderName $name) -and
            ($name -ine "Dups")
        ) {
            $mistakes.Add(
                "FOLDER NAME should be 'Dups': $relativeFolder"
            )
        }

        if (
            (Test-EFormsFolderName $name) -and
            ($name -ine "E-forms")
        ) {
            $mistakes.Add(
                "FOLDER NAME should be 'E-forms': $relativeFolder"
            )
        }

        if (
            (Test-PaperFormsFolderName $name) -and
            ($name -ine "Paper forms")
        ) {
            $mistakes.Add(
                "FOLDER NAME should be 'Paper forms': $relativeFolder"
            )
        }

        if (
            (Test-EnteredD4wFolderName $name) -and
            ($name -ine "Entered in D4w")
        ) {
            $mistakes.Add(
                "FOLDER NAME should be 'Entered in D4w': $relativeFolder"
            )
        }

        if (
            (Test-FormsFromT1FolderName $name) -and
            ($name -ine "Forms from t1")
        ) {
            $mistakes.Add(
                "FOLDER NAME should be 'Forms from t1': $relativeFolder"
            )
        }
    }


    # ---------- Consent folders = one paper form ----------

    $consentFolders = New-Object System.Collections.Generic.List[object]

    $possibleConsentFolders = @(
        $allDirectories |
        Where-Object {
            Test-ConsentFormFolderName $_.Name
        } |
        Sort-Object {
            $_.FullName.Length
        }
    )

    foreach ($directory in $possibleConsentFolders) {

        $nestedInsideExistingConsent = $false

        foreach ($existing in $consentFolders) {

            $prefix = $existing.FullName.TrimEnd('\') + "\"

            if (
                $directory.FullName.StartsWith(
                    $prefix,
                    [System.StringComparison]::OrdinalIgnoreCase
                )
            ) {
                $nestedInsideExistingConsent = $true
                break
            }
        }

        if (-not $nestedInsideExistingConsent) {
            $consentFolders.Add($directory)
        }
    }


    # ---------- Count/check Consent folders ----------

    foreach ($consentFolder in $consentFolders) {

        # A Consent folder represents ONE paper form.
        # It is always included in the displayed Paper forms number.
        $paperForms++

        $relativePath = $consentFolder.FullName.Substring(
            $SchoolPath.Length
        ).TrimStart('\')

        $flags = Get-LocationFlags -RelativePath $relativePath

        # Paper forms inside Dups are shown in Paper forms,
        # but are NOT included in Total forms.
        if (-not $flags.InsideDups) {
            $paperFormsForTotal++
        }

        # A Consent folder is loose when it is outside all workflow
        # and duplicate/category folders. It still counts as one paper form.
        $isLooseConsentFolder = (
            -not $flags.InsideDups -and
            (
                $flags.InsideLate -or
                (
                    -not $flags.InsideEntered -and
                    -not $flags.InsideEForms -and
                    -not $flags.InsidePaperForms
                )
            )
        )

        if ($isLooseConsentFolder) {

            if ($flags.InsideFormsFromT1) {
                $formsFromT1Loose++
            }
            else {
                $looseForms++
            }
        }

        $directlyInsideEntered = $false

        if ($null -ne $consentFolder.Parent) {
            $directlyInsideEntered = (
                Test-EnteredD4wFolderName $consentFolder.Parent.Name
            )
        }


        if ($flags.InsideLate) {
            # Late forms are unworked, regardless of their actual form type.
            # Count them normally but do not report placement errors.
        }

        elseif (
            $flags.InsideEForms -and
            $flags.InsidePaperForms
        ) {
            $mistakes.Add(
                "PAPER FORM folder path contains BOTH E-forms and Paper forms: $relativePath"
            )
        }

        elseif ($flags.InsideEForms) {
            $mistakes.Add(
                "PAPER FORM folder located inside E-FORMS folder - MOVE to Paper forms: $relativePath"
            )
        }

        elseif (
            $directlyInsideEntered -and
            -not $flags.InsideDups
        ) {
            $mistakes.Add(
                "UNSEPARATED PAPER FORM folder directly inside 'Entered in D4w' - MOVE to Paper forms: $relativePath"
            )

            $unseparated.Add($relativePath)
        }

        elseif (
            $flags.InsideEntered -and
            -not $flags.InsidePaperForms -and
            -not $flags.InsideDups
        ) {
            $mistakes.Add(
                "PAPER FORM folder inside 'Entered in D4w' but NOT inside Paper forms - MOVE to Paper forms: $relativePath"
            )
        }

        elseif (
            $flags.InsidePaperForms -and
            -not $flags.InsideEntered -and
            -not $flags.InsideDups
        ) {
            $mistakes.Add(
                "PAPER FORM folder is inside Paper forms OUTSIDE 'Entered in D4w': $relativePath"
            )
        }
    }


    # ---------- Read PDFs ----------

    try {
        $pdfFiles = @(
            Get-ChildItem `
                -LiteralPath $SchoolPath `
                -Filter "*.pdf" `
                -File `
                -Recurse `
                -ErrorAction Stop
        )
    }
    catch {
        $mistakes.Add(
            "COULD NOT ENUMERATE PDF FILES: $($_.Exception.Message)"
        )
        $pdfFiles = @()
    }


    # ---------- Scan PDFs ----------

    foreach ($file in $pdfFiles) {

        try {

            # PDFs inside a Consent folder are NOT counted
            # separately. The folder already counted as 1 paper form.
            $insideConsentFolder = $false

            foreach ($consentFolder in $consentFolders) {

                $prefix = $consentFolder.FullName.TrimEnd('\') + "\"

                if (
                    $file.FullName.StartsWith(
                        $prefix,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                ) {
                    $insideConsentFolder = $true
                    break
                }
            }

            if ($insideConsentFolder) {
                continue
            }


            $pages = Get-PdfPageCount -Path $file.FullName
            $sizeBytes = $file.Length
            $sizeKB = [math]::Round($sizeBytes / 1KB, 1)

            $relativePath = $file.FullName.Substring(
                $SchoolPath.Length
            ).TrimStart('\')

            $relativeDirectory = $file.DirectoryName.Substring(
                $SchoolPath.Length
            ).TrimStart('\')

            $flags = Get-LocationFlags `
                -RelativePath $relativeDirectory

            # Loose/unworked form:
            # outside Entered in D4w, Dups, E-forms and Paper forms.
            # It still counts normally, but naming/workflow checks
            # are not enforced yet.
            $isLoose = (
                -not $flags.InsideDups -and
                (
                    $flags.InsideLate -or
                    (
                        -not $flags.InsideEntered -and
                        -not $flags.InsideEForms -and
                        -not $flags.InsidePaperForms
                    )
                )
            )

            # This is the exact condition the new report cares about:
            # the PDF itself is sitting DIRECTLY inside Entered in D4w.
            $directlyInsideEntered = (
                Test-EnteredD4wFolderName $file.Directory.Name
            )

            $hasCdbs = (
                $file.BaseName -match '(?i)cdbs'
            )

            $hasConsentForm = (
                $file.BaseName -match '(?i)consent\s*form'
            )


            # ---------- Identify actual type ----------

            $isEForm = $false
            $isPaperForm = $false
            $formType = ""


            # CDBS E-form
            if (
                $pages -eq 2 -and
                $hasCdbs
            ) {
                $isEForm = $true
                $formType = "CDBS E-FORM"

                if ($sizeBytes -ge 200KB) {
                    $mistakes.Add(
                        "CDBS E-FORM unusual size: $relativePath [$sizeKB KB]"
                    )
                }
            }


            # Paper form
            elseif ($pages -eq 2) {
                $isPaperForm = $true
                $formType = "PAPER FORM"

                if (
                    -not $isLoose -and
                    -not $hasConsentForm
                ) {
                    $mistakes.Add(
                        "PAPER FORM missing 'Consent Form' in name: $relativePath [$sizeKB KB]"
                    )
                }
            }


            # Old E-form
            elseif (
                (
                    $pages -eq 5 -or
                    $pages -eq 6
                ) -and
                $sizeBytes -lt 100KB
            ) {
                $isEForm = $true
                $formType = "OLD E-FORM"

                if (
                    -not $isLoose -and
                    -not $hasConsentForm
                ) {
                    $mistakes.Add(
                        "OLD E-FORM missing 'Consent Form' in name: $relativePath [$pages pages, $sizeKB KB]"
                    )
                }
            }


            # New E-form
            elseif (
                $pages -eq 7 -or
                $pages -eq 8
            ) {
                $isEForm = $true
                $formType = "NEW E-FORM"

                if (
                    -not $isLoose -and
                    $hasConsentForm
                ) {
                    $mistakes.Add(
                        "NEW E-FORM contains 'Consent Form' in name: $relativePath [$pages pages, $sizeKB KB]"
                    )
                }
            }


            # Other PDF over 2 pages
            elseif ($pages -gt 2) {
                $isEForm = $true
                $formType = "E-FORM"

                $mistakes.Add(
                    "UNUSUAL E-FORM pattern: $relativePath [$pages pages, $sizeKB KB]"
                )
            }


            # 0/1-page PDF outside a Consent folder
            else {
                $mistakes.Add(
                    "UNEXPECTED PAGE COUNT: $relativePath [$pages page(s), $sizeKB KB]"
                )
            }


            # ---------- Loose counts ----------
            #
            # Loose forms count normally in Total / E-forms / Paper forms.
            # These two counters only show where the unworked forms are.

            if (
                $isLoose -and
                (
                    $isEForm -or
                    $isPaperForm
                )
            ) {

                if ($flags.InsideFormsFromT1) {
                    $formsFromT1Loose++
                }
                else {
                    $looseForms++
                }
            }


            # ---------- Count ----------

            if ($isPaperForm) {

                # Displayed Paper forms includes paper forms everywhere,
                # including inside ANY Dups folder.
                $paperForms++

                # Total forms excludes paper forms inside Dups.
                if (-not $flags.InsideDups) {
                    $paperFormsForTotal++
                }
            }

            elseif ($isEForm) {
                # E-forms inside Dups do NOT count.
                if (-not $flags.InsideDups) {
                    $eForms++
                }
            }


            # ---------- Physical location checks ----------

            if ($flags.InsideLate) {
                # Do not mark Late forms as unseparated or misplaced.
                # Actual type was already identified from the PDF above.
            }

            elseif (
                $flags.InsideEForms -and
                $flags.InsidePaperForms
            ) {
                $mistakes.Add(
                    "FORM path contains BOTH E-forms and Paper forms folders: $relativePath"
                )
            }

            elseif ($flags.InsideEntered) {

                # Paper inside E-forms
                if (
                    $isPaperForm -and
                    $flags.InsideEForms
                ) {
                    $mistakes.Add(
                        "PAPER FORM located inside E-FORMS folder - MOVE to Paper forms: $relativePath"
                    )
                }


                # E-form inside Paper forms
                elseif (
                    $isEForm -and
                    $flags.InsidePaperForms
                ) {
                    $mistakes.Add(
                        "$formType located inside PAPER FORMS folder - MOVE to E-forms: $relativePath"
                    )
                }


                # NEW FEATURE:
                # File left DIRECTLY in Entered in D4w.
                elseif (
                    $directlyInsideEntered -and
                    -not $flags.InsideDups -and
                    (
                        $isEForm -or
                        $isPaperForm
                    )
                ) {

                    if ($isPaperForm) {
                        $mistakes.Add(
                            "UNSEPARATED PAPER FORM directly inside 'Entered in D4w' - MOVE to Paper forms: $relativePath"
                        )
                    }
                    else {
                        $mistakes.Add(
                            "UNSEPARATED $formType directly inside 'Entered in D4w' - MOVE to E-forms: $relativePath"
                        )
                    }

                    $unseparated.Add($relativePath)
                }


                # Other unsorted location somewhere under Entered
                elseif (
                    -not $flags.InsideDups -and
                    -not $flags.InsideEForms -and
                    -not $flags.InsidePaperForms -and
                    (
                        $isEForm -or
                        $isPaperForm
                    )
                ) {

                    if ($isPaperForm) {
                        $mistakes.Add(
                            "PAPER FORM inside 'Entered in D4w' but NOT inside Paper forms - MOVE to Paper forms: $relativePath"
                        )
                    }
                    else {
                        $mistakes.Add(
                            "$formType inside 'Entered in D4w' but NOT inside E-forms - MOVE to E-forms: $relativePath"
                        )
                    }
                }
            }

            elseif (
                (
                    $flags.InsideEForms -or
                    $flags.InsidePaperForms
                ) -and
                (
                    $isEForm -or
                    $isPaperForm
                )
            ) {
                $mistakes.Add(
                    "FORM is inside an E-forms/Paper forms folder OUTSIDE 'Entered in D4w': $relativePath"
                )
            }
        }

        catch {

            $relativePath = $file.FullName.Substring(
                $SchoolPath.Length
            ).TrimStart('\')

            $mistakes.Add(
                "COULD NOT READ PDF: $relativePath"
            )
        }
    }


    # Total forms excludes ALL forms inside Dups.
    #
    # E-forms already exclude Dups when counted.
    # PaperFormsForTotal excludes paper forms inside Dups.
    #
    # The displayed PaperForms value still includes papers in Dups.
    $totalForms = $eForms + $paperFormsForTotal


    return [PSCustomObject]@{
        SchoolName       = $schoolName
        SchoolPath       = $SchoolPath
        TotalForms       = $totalForms
        EForms           = $eForms
        PaperForms        = $paperForms
        LooseForms        = $looseForms
        FormsFromT1Loose  = $formsFromT1Loose
        HasFormsFromT1    = $hasFormsFromT1
        MistakeCount      = $mistakes.Count
        Mistakes          = $mistakes
        UnseparatedCount  = $unseparated.Count
        UnseparatedForms  = $unseparated
    }
}


# ============================================================
# CHOOSE MODE
# ============================================================

$modeChoice = [System.Windows.Forms.MessageBox]::Show(
    "YES = Scan ALL schools`n`nNO = Scan ONE school`n`nCANCEL = Exit",
    "School Form Counter",
    [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if (
    $modeChoice -eq
    [System.Windows.Forms.DialogResult]::Cancel
) {
    return
}

$batchMode = (
    $modeChoice -eq
    [System.Windows.Forms.DialogResult]::Yes
)


# ============================================================
# SHORTCUT OR NORMAL BROWSING
# ============================================================

$pathChoice = [System.Windows.Forms.MessageBox]::Show(
    "YES = Use the Schools shortcut from your Desktop`n`nNO = Browse normally`n`nCANCEL = Exit",
    "Choose Folder",
    [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if (
    $pathChoice -eq
    [System.Windows.Forms.DialogResult]::Cancel
) {
    return
}

$useShortcut = (
    $pathChoice -eq
    [System.Windows.Forms.DialogResult]::Yes
)


# ============================================================
# RESOLVE SHORTCUT
# ============================================================

$shortcutRoot = $null

if ($useShortcut) {

    $shortcutPicker = New-Object System.Windows.Forms.OpenFileDialog

    $shortcutPicker.Title = "Select the Schools folder shortcut"

    $shortcutPicker.Filter = "Windows shortcuts (*.lnk)|*.lnk"

    $shortcutPicker.Multiselect = $false

    $shortcutPicker.InitialDirectory = (
        [Environment]::GetFolderPath("Desktop")
    )

    if (
        $shortcutPicker.ShowDialog() -ne
        [System.Windows.Forms.DialogResult]::OK
    ) {
        return
    }

    try {

        $shell = New-Object -ComObject WScript.Shell

        $shortcut = $shell.CreateShortcut(
            $shortcutPicker.FileName
        )

        $shortcutRoot = $shortcut.TargetPath

        if (
            [string]::IsNullOrWhiteSpace($shortcutRoot) -or
            -not (
                Test-Path `
                    -LiteralPath $shortcutRoot `
                    -PathType Container
            )
        ) {
            [System.Windows.Forms.MessageBox]::Show(
                "The shortcut does not point to a valid folder.",
                "Invalid Shortcut",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }
    }

    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not read the shortcut.`n`n$($_.Exception.Message)",
            "Shortcut Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
}


# ============================================================
# BATCH MODE
# ============================================================

if ($batchMode) {

    if ($useShortcut) {
        $schoolsRoot = $shortcutRoot
    }
    else {
        $rootPicker = New-Object System.Windows.Forms.FolderBrowserDialog

        $rootPicker.Description = "Select the folder containing ALL school folders"

        $rootPicker.ShowNewFolderButton = $false

        if (
            $rootPicker.ShowDialog() -ne
            [System.Windows.Forms.DialogResult]::OK
        ) {
            return
        }

        $schoolsRoot = $rootPicker.SelectedPath
    }


    # Immediate subfolders = schools
    try {
        $schoolDirectories = @(
            Get-ChildItem `
                -LiteralPath $schoolsRoot `
                -Directory `
                -ErrorAction Stop |
            Sort-Object Name
        )
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not read the school folders.`n`n$($_.Exception.Message)",
            "Folder Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }


    if ($schoolDirectories.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No school folders were found inside:`n`n$schoolsRoot",
            "Nothing Found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }


    $results = New-Object System.Collections.Generic.List[object]

    $schoolNumber = 0

    foreach ($school in $schoolDirectories) {

        $schoolNumber++

        $percent = [int](
            ($schoolNumber / $schoolDirectories.Count) * 100
        )

        Write-Progress `
            -Activity "Scanning school folders" `
            -Status "$schoolNumber of $($schoolDirectories.Count): $($school.Name)" `
            -PercentComplete $percent

        $result = Scan-School `
            -SchoolPath $school.FullName

        $results.Add($result)
    }

    Write-Progress `
        -Activity "Scanning school folders" `
        -Completed


    # ---------- Overall totals ----------

    $allTotalForms = 0
    $allEForms = 0
    $allPaperForms = 0
    $allLooseForms = 0
    $allFormsFromT1Loose = 0
    $schoolsWithoutFormsFromT1 = 0
    $allMistakes = 0
    $allUnseparated = 0

    foreach ($result in $results) {
        $allTotalForms += $result.TotalForms
        $allEForms += $result.EForms
        $allPaperForms += $result.PaperForms
        $allLooseForms += $result.LooseForms
        $allFormsFromT1Loose += $result.FormsFromT1Loose

        if (-not $result.HasFormsFromT1) {
            $schoolsWithoutFormsFromT1++
        }

        $allMistakes += $result.MistakeCount
        $allUnseparated += $result.UnseparatedCount
    }


    $schoolsWithUnseparated = @(
        $results |
        Where-Object {
            $_.UnseparatedCount -gt 0
        } |
        Sort-Object SchoolName
    )

    $schoolsWithMistakes = @(
        $results |
        Where-Object {
            $_.MistakeCount -gt 0
        } |
        Sort-Object SchoolName
    )


    # ---------- Build report ----------

    $report = New-Object System.Collections.Generic.List[string]

    $report.Add("============================================================")
    $report.Add("SCHOOL FORMS BATCH REPORT")
    $report.Add("============================================================")
    $report.Add("")
    $report.Add("Parent folder: $schoolsRoot")
    $report.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $report.Add("")
    $report.Add("Schools scanned: $($results.Count)")
    $report.Add("")

    $report.Add("============================================================")
    $report.Add("OVERALL TOTALS")
    $report.Add("============================================================")
    $report.Add("")
    $report.Add("Total forms:       $allTotalForms")
    $report.Add("E-forms:           $allEForms")
    $report.Add("Paper forms:                $allPaperForms")
    $report.Add("Loose forms:                $allLooseForms")
    $report.Add("Forms from t1 loose:        $allFormsFromT1Loose")
    $report.Add("Schools without From t1:    $schoolsWithoutFormsFromT1")
    $report.Add("Total mistakes:             $allMistakes")
    $report.Add("Unseparated forms:          $allUnseparated")
    $report.Add("")


    # ---------- Loose forms summary ----------

    $report.Add("============================================================")
    $report.Add("LOOSE FORMS BY SCHOOL")
    $report.Add("============================================================")
    $report.Add("")

    foreach ($result in $results) {

        $report.Add($result.SchoolName)
        $report.Add("  Loose forms: $($result.LooseForms)")

        if ($result.HasFormsFromT1) {
            $report.Add(
                "  Forms from t1 loose: $($result.FormsFromT1Loose)"
            )
        }
        else {
            $report.Add(
                "  Forms from t1 loose: NO 'Forms from t1' FOLDER"
            )
        }

        $report.Add("")
    }

    $report.Add("")


    # ---------- Special unseparated summary ----------

    $report.Add("============================================================")
    $report.Add("SCHOOLS WITH UNSEPARATED FORMS IN ENTERED IN D4W")
    $report.Add("============================================================")
    $report.Add("")

    if ($schoolsWithUnseparated.Count -eq 0) {

        $report.Add(
            "None. No forms were found sitting directly inside Entered in D4w."
        )
    }
    else {

        foreach ($result in $schoolsWithUnseparated) {

            $report.Add(
                "$($result.SchoolName): $($result.UnseparatedCount) unseparated form(s)"
            )

            foreach ($item in $result.UnseparatedForms) {
                $report.Add("  $item")
            }

            $report.Add("")
        }
    }

    $report.Add("")


    # ---------- Schools with any mistakes ----------

    $report.Add("============================================================")
    $report.Add("SCHOOLS WITH MISTAKES")
    $report.Add("============================================================")
    $report.Add("")

    if ($schoolsWithMistakes.Count -eq 0) {
        $report.Add("None. No mistakes found in any school.")
    }
    else {
        foreach ($result in $schoolsWithMistakes) {
            $report.Add(
                "$($result.SchoolName): $($result.MistakeCount) mistake(s)"
            )
        }
    }

    $report.Add("")


    # ---------- Detailed results ----------

    foreach ($result in $results) {

        $report.Add("")
        $report.Add("============================================================")
        $report.Add($result.SchoolName)
        $report.Add("============================================================")
        $report.Add("")
        $report.Add("Path: $($result.SchoolPath)")
        $report.Add("")
        $report.Add("Total forms:            $($result.TotalForms)")
        $report.Add("E-forms:                $($result.EForms)")
        $report.Add("Paper forms:            $($result.PaperForms)")
        $report.Add("Loose forms:            $($result.LooseForms)")

        if ($result.HasFormsFromT1) {
            $report.Add(
                "Forms from t1 loose:      $($result.FormsFromT1Loose)"
            )
        }
        else {
            $report.Add(
                "Forms from t1 loose:      NO 'Forms from t1' FOLDER"
            )
        }

        $report.Add("Unseparated in Entered: $($result.UnseparatedCount)")
        $report.Add("")
        $report.Add("Mistakes:               $($result.MistakeCount)")
        $report.Add("")

        if ($result.UnseparatedCount -gt 0) {

            $report.Add("UNSEPARATED FORMS:")

            foreach ($item in $result.UnseparatedForms) {
                $report.Add("  $item")
            }

            $report.Add("")
        }

        if ($result.MistakeCount -eq 0) {
            $report.Add("No mistakes found.")
        }
        else {
            $report.Add("ALL MISTAKES:")

            foreach ($mistake in $result.Mistakes) {
                $report.Add("  $mistake")
            }
        }

        $report.Add("")
    }


    # ---------- Save ONE report to Desktop ----------

    $desktop = [Environment]::GetFolderPath("Desktop")
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

    $reportPath = Join-Path `
        $desktop `
        "School Form Batch Report $timestamp.txt"

    # This is the ONLY write performed by batch mode.
    $report |
        Out-File `
            -LiteralPath $reportPath `
            -Encoding UTF8


    # ---------- Console summary ----------

    Clear-Host

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " SCHOOL FORMS BATCH COMPLETE" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Schools scanned:   " -NoNewline
    Write-Host $results.Count -ForegroundColor Green

    Write-Host ""
    Write-Host "Total forms:       " -NoNewline
    Write-Host $allTotalForms -ForegroundColor Green

    Write-Host "E-forms:           " -NoNewline
    Write-Host $allEForms -ForegroundColor Green

    Write-Host "Paper forms:       " -NoNewline
    Write-Host $allPaperForms -ForegroundColor Green

    Write-Host "Loose forms:       " -NoNewline
    Write-Host $allLooseForms -ForegroundColor Cyan

    Write-Host "From t1 loose:     " -NoNewline
    Write-Host $allFormsFromT1Loose -ForegroundColor Cyan

    Write-Host "No From t1 folder: " -NoNewline
    Write-Host $schoolsWithoutFormsFromT1 -ForegroundColor Cyan

    Write-Host ""
    Write-Host "Total mistakes:    " -NoNewline

    if ($allMistakes -eq 0) {
        Write-Host "0" -ForegroundColor Green
    }
    else {
        Write-Host $allMistakes -ForegroundColor Yellow
    }

    Write-Host "Unseparated forms: " -NoNewline

    if ($allUnseparated -eq 0) {
        Write-Host "0" -ForegroundColor Green
    }
    else {
        Write-Host $allUnseparated -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "Report saved to:" -ForegroundColor Cyan
    Write-Host $reportPath
    Write-Host ""


    if ($schoolsWithUnseparated.Count -gt 0) {

        Write-Host "Schools with UNSEPARATED forms:" -ForegroundColor Red

        foreach ($result in $schoolsWithUnseparated) {
            Write-Host (
                "  {0}: {1}" -f
                $result.SchoolName,
                $result.UnseparatedCount
            ) -ForegroundColor Red
        }

        Write-Host ""
    }


    if ($schoolsWithMistakes.Count -gt 0) {

        Write-Host "Schools with mistakes:" -ForegroundColor Yellow

        foreach ($result in $schoolsWithMistakes) {
            Write-Host (
                "  {0}: {1}" -f
                $result.SchoolName,
                $result.MistakeCount
            ) -ForegroundColor Yellow
        }

        Write-Host ""
    }

    return
}


# ============================================================
# SINGLE SCHOOL MODE
# ============================================================

else {

    if ($useShortcut) {

        $schoolsRoot = $shortcutRoot

        $rootNormalized = (
            [System.IO.Path]::GetFullPath($schoolsRoot)
        ).TrimEnd('\')

        while ($true) {

            $schoolPicker = New-Object System.Windows.Forms.FolderBrowserDialog

            $schoolPicker.Description = "Select the school you want to scan"

            $schoolPicker.ShowNewFolderButton = $false

            $schoolPicker.SelectedPath = $schoolsRoot

            if (
                $schoolPicker.ShowDialog() -ne
                [System.Windows.Forms.DialogResult]::OK
            ) {
                return
            }

            $selectedSchool = $schoolPicker.SelectedPath

            $selectedNormalized = (
                [System.IO.Path]::GetFullPath($selectedSchool)
            ).TrimEnd('\')

            if ($selectedNormalized -ieq $rootNormalized) {

                [System.Windows.Forms.MessageBox]::Show(
                    "Select an individual school inside the Schools folder.`n`nDo not select the Schools folder itself.",
                    "Select One School",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )

                continue
            }

            $insideRoot = $selectedNormalized.StartsWith(
                $rootNormalized + "\",
                [System.StringComparison]::OrdinalIgnoreCase
            )

            if (-not $insideRoot) {

                [System.Windows.Forms.MessageBox]::Show(
                    "Please select a school inside:`n`n$schoolsRoot",
                    "Wrong Folder",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )

                continue
            }

            $schoolPath = $selectedSchool
            break
        }
    }

    else {

        $schoolPicker = New-Object System.Windows.Forms.FolderBrowserDialog

        $schoolPicker.Description = "Select the TOP-LEVEL school folder"

        $schoolPicker.ShowNewFolderButton = $false

        if (
            $schoolPicker.ShowDialog() -ne
            [System.Windows.Forms.DialogResult]::OK
        ) {
            return
        }

        $schoolPath = $schoolPicker.SelectedPath
    }


    $result = Scan-School -SchoolPath $schoolPath


    Clear-Host

    Write-Host ""
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host " SCHOOL FORM COUNT" -ForegroundColor Cyan
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "School: $($result.SchoolName)"
    Write-Host "Path:   $($result.SchoolPath)"
    Write-Host ""

    Write-Host "Total forms:            " -NoNewline
    Write-Host $result.TotalForms -ForegroundColor Green

    Write-Host "E-forms:                " -NoNewline
    Write-Host $result.EForms -ForegroundColor Green

    Write-Host "Paper forms:            " -NoNewline
    Write-Host $result.PaperForms -ForegroundColor Green

    Write-Host "Loose forms:            " -NoNewline
    Write-Host $result.LooseForms -ForegroundColor Cyan

    Write-Host "Forms from t1 loose:    " -NoNewline

    if ($result.HasFormsFromT1) {
        Write-Host $result.FormsFromT1Loose -ForegroundColor Cyan
    }
    else {
        Write-Host "NO 'Forms from t1' FOLDER" -ForegroundColor Yellow
    }

    Write-Host "Unseparated in Entered: " -NoNewline

    if ($result.UnseparatedCount -eq 0) {
        Write-Host "0" -ForegroundColor Green
    }
    else {
        Write-Host $result.UnseparatedCount -ForegroundColor Red
    }

    Write-Host ""

    if ($result.UnseparatedCount -gt 0) {

        Write-Host "UNSEPARATED FORMS:" -ForegroundColor Red

        foreach ($item in $result.UnseparatedForms) {
            Write-Host "  $item" -ForegroundColor Red
        }

        Write-Host ""
    }

    Write-Host "Mistakes:               " -NoNewline

    if ($result.MistakeCount -eq 0) {
        Write-Host "0" -ForegroundColor Green
        Write-Host ""
        Write-Host "No mistakes found." -ForegroundColor Green
    }
    else {
        Write-Host $result.MistakeCount -ForegroundColor Yellow
        Write-Host ""

        foreach ($mistake in $result.Mistakes) {
            Write-Host "  $mistake" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host ""
}
