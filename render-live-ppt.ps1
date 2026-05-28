<# 
Renders a JSON slide spec into a native PowerPoint deck through the PowerPoint
COM object model. It intentionally draws element-by-element so the user can see
PowerPoint being operated in real time.

Requirements:
- Windows desktop PowerPoint installed.
- Run from PowerShell on the same desktop session as the user.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Spec,

    [string]$Output,

    [bool]$Visible = $true,

    [bool]$KeepOpen = $true,

    [switch]$CloseAfterSave,

    [int]$StepDelayMs = 120,

    [switch]$ValidateOnly,

    [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$MsoTrue = -1
$MsoFalse = 0
$PpLayoutBlank = 12
$PpSaveAsOpenXmlPresentation = 24
$MsoTextOrientationHorizontal = 1
$MsoScaleFromTopLeft = 0

function Get-Value {
    param(
        [object]$Object,
        [string]$Name,
        [object]$Default = $null
    )

    if ($null -eq $Object) {
        return $Default
    }

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name) -and $null -ne $Object[$Name]) {
            return $Object[$Name]
        }
        return $Default
    }

    $prop = $Object.PSObject.Properties[$Name]
    if ($null -ne $prop -and $null -ne $prop.Value) {
        return $prop.Value
    }

    return $Default
}

function ConvertTo-PptColor {
    param(
        [object]$Value,
        [object]$Default = "#000000"
    )

    if ($null -eq $Value -or [string]$Value -eq "") {
        $Value = $Default
    }

    if ($Value -is [int]) {
        return [int]$Value
    }

    $hex = ([string]$Value).Trim()
    if ($hex.StartsWith("0x")) {
        return [Convert]::ToInt32($hex.Substring(2), 16)
    }

    $hex = $hex.TrimStart("#")
    if ($hex.Length -eq 3) {
        $hex = -join ($hex.ToCharArray() | ForEach-Object { "$_$_" })
    }
    if ($hex.Length -ne 6) {
        throw "Invalid color value '$Value'. Use #RRGGBB."
    }

    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return ($r + ($g -shl 8) + ($b -shl 16))
}

function Get-Number {
    param(
        [object]$Value,
        [double]$Default = 0
    )

    if ($null -eq $Value -or [string]$Value -eq "") {
        return $Default
    }
    return [double]$Value
}

function Get-BoolAsMso {
    param(
        [object]$Value,
        [bool]$Default = $false
    )

    if ($null -eq $Value) {
        $Value = $Default
    }
    if ([bool]$Value) {
        return $script:MsoTrue
    }
    return $script:MsoFalse
}

function Get-ShapeType {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = "roundRect"
    }
    switch ($Name.ToLowerInvariant()) {
        "rect" { return 1 }
        "rectangle" { return 1 }
        "roundrect" { return 5 }
        "roundedrectangle" { return 5 }
        "oval" { return 9 }
        "circle" { return 9 }
        "diamond" { return 4 }
        "triangle" { return 7 }
        "hexagon" { return 10 }
        "parallelogram" { return 2 }
        default { throw "Unsupported shape '$Name'." }
    }
}

function Get-ConnectorType {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = "straight"
    }
    switch ($Name.ToLowerInvariant()) {
        "straight" { return 1 }
        "elbow" { return 2 }
        "curve" { return 3 }
        default { throw "Unsupported connector '$Name'." }
    }
}

function Get-ArrowheadStyle {
    param([object]$Value)

    if ($null -eq $Value -or [string]$Value -eq "" -or [string]$Value -eq "none") {
        return 1
    }

    switch (([string]$Value).ToLowerInvariant()) {
        "triangle" { return 3 }
        "stealth" { return 4 }
        "oval" { return 5 }
        "diamond" { return 6 }
        default { return 3 }
    }
}

function Get-TextAlignment {
    param([object]$Value)

    if ($null -eq $Value -or [string]$Value -eq "") {
        $Value = "center"
    }
    switch ((([string]$Value).ToLowerInvariant())) {
        "left" { return 1 }
        "center" { return 2 }
        "right" { return 3 }
        default { return 2 }
    }
}

function Get-VerticalAnchor {
    param([object]$Value)

    if ($null -eq $Value -or [string]$Value -eq "") {
        $Value = "middle"
    }
    switch ((([string]$Value).ToLowerInvariant())) {
        "top" { return 1 }
        "middle" { return 3 }
        "center" { return 3 }
        "bottom" { return 4 }
        default { return 3 }
    }
}

function Resolve-ElementPath {
    param(
        [string]$Path,
        [string]$BaseDir
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }
    return [System.IO.Path]::GetFullPath((Join-Path $BaseDir $Path))
}

function Build-AssetMap {
    param(
        [object]$DeckSpec,
        [string]$SpecDir
    )

    $map = @{}
    $assetsRaw = Get-Value $DeckSpec "assets" @()
    $assetItems = @()

    if ($null -eq $assetsRaw) {
        return $map
    }
    elseif ($assetsRaw -is [System.Array]) {
        $assetItems = @($assetsRaw)
    }
    elseif ($assetsRaw -is [System.Collections.IDictionary]) {
        foreach ($key in $assetsRaw.Keys) {
            $asset = $assetsRaw[$key]
            if ($null -eq (Get-Value $asset "id" $null)) {
                $asset | Add-Member -NotePropertyName "id" -NotePropertyValue $key -Force
            }
            $assetItems += $asset
        }
    }
    elseif ($null -ne $assetsRaw.PSObject.Properties["id"]) {
        $assetItems = @($assetsRaw)
    }
    else {
        foreach ($prop in $assetsRaw.PSObject.Properties) {
            $asset = $prop.Value
            if ($null -eq (Get-Value $asset "id" $null)) {
                $asset | Add-Member -NotePropertyName "id" -NotePropertyValue $prop.Name -Force
            }
            $assetItems += $asset
        }
    }

    foreach ($asset in @($assetItems)) {
        $id = Get-Value $asset "id" $null
        if ($null -eq $id -or [string]$id -eq "") {
            throw "Every asset must have an id."
        }

        $path = Get-Value $asset "path" (Get-Value $asset "file" $null)
        if ($null -ne $path -and [string]$path -ne "") {
            $resolved = Resolve-ElementPath ([string]$path) $SpecDir
            $asset | Add-Member -NotePropertyName "resolvedPath" -NotePropertyValue $resolved -Force
        }

        $map[[string]$id] = $asset
    }

    return $map
}

function Resolve-ImagePath {
    param(
        [object]$Element,
        [string]$SpecDir,
        [hashtable]$AssetById
    )

    $assetId = Get-Value $Element "asset" $null
    if ($null -ne $assetId -and [string]$assetId -ne "") {
        if (-not $AssetById.ContainsKey([string]$assetId)) {
            throw "Unknown image asset '$assetId'. Add it to top-level assets."
        }

        $asset = $AssetById[[string]$assetId]
        $resolved = Get-Value $asset "resolvedPath" $null
        if ($null -eq $resolved -or [string]$resolved -eq "") {
            $prompt = Get-Value $asset "prompt" ""
            if ([string]$prompt -ne "") {
                throw "Asset '$assetId' has a prompt but no path. Generate it with Codex image generation, save it locally, and set assets[].path."
            }
            throw "Asset '$assetId' has no path."
        }
        return [string]$resolved
    }

    $path = Get-Value $Element "path" $null
    if ($null -eq $path -or [string]$path -eq "") {
        $prompt = Get-Value $Element "prompt" ""
        if ([string]$prompt -ne "") {
            throw "Image element has a prompt but no path. Generate the image first and reference the saved file."
        }
        throw "Image element requires either 'path' or 'asset'."
    }

    return Resolve-ElementPath ([string]$path) $SpecDir
}

function Wait-DrawStep {
    if ($script:StepDelayMs -gt 0) {
        Start-Sleep -Milliseconds $script:StepDelayMs
    }
}

function Set-FillStyle {
    param(
        [object]$Shape,
        [object]$Fill,
        [object]$DefaultColor = "#FFFFFF"
    )

    if ($null -eq $Fill) {
        $Fill = $DefaultColor
    }

    if ([string]$Fill -eq "none") {
        $Shape.Fill.Visible = $script:MsoFalse
        return
    }

    $color = $Fill
    $transparency = 0
    if ($Fill -isnot [string] -and $Fill -isnot [int]) {
        $color = Get-Value $Fill "color" $DefaultColor
        $transparency = Get-Number (Get-Value $Fill "transparency" 0) 0
    }

    $Shape.Fill.Visible = $script:MsoTrue
    $Shape.Fill.ForeColor.RGB = ConvertTo-PptColor $color $DefaultColor
    $Shape.Fill.Transparency = $transparency
}

function Set-LineStyle {
    param(
        [object]$Shape,
        [object]$Line,
        [object]$DefaultColor = "#C9D3E0"
    )

    if ($null -eq $Line) {
        $Line = @{ color = $DefaultColor; width = 1 }
    }

    if ([string]$Line -eq "none") {
        $Shape.Line.Visible = $script:MsoFalse
        return
    }

    $color = Get-Value $Line "color" $DefaultColor
    $width = Get-Number (Get-Value $Line "width" 1) 1
    $beginArrow = Get-Value $Line "beginArrow" "none"
    $endArrow = Get-Value $Line "endArrow" "none"

    $Shape.Line.Visible = $script:MsoTrue
    $Shape.Line.ForeColor.RGB = ConvertTo-PptColor $color $DefaultColor
    $Shape.Line.Weight = $width
    if ($null -ne $beginArrow -and [string]$beginArrow -ne "" -and [string]$beginArrow -ne "none") {
        $Shape.Line.BeginArrowheadStyle = Get-ArrowheadStyle $beginArrow
    }
    if ($null -ne $endArrow -and [string]$endArrow -ne "" -and [string]$endArrow -ne "none") {
        $Shape.Line.EndArrowheadStyle = Get-ArrowheadStyle $endArrow
    }
}

function Set-TextStyle {
    param(
        [object]$Shape,
        [object]$TextSpec,
        [object]$Theme
    )

    if ($null -eq $TextSpec) {
        return
    }

    $text = $TextSpec
    if ($TextSpec -isnot [string]) {
        $text = Get-Value $TextSpec "text" ""
    }
    if ([string]$text -eq "") {
        return
    }

    $font = Get-Value $TextSpec "font" (Get-Value $Theme "font" "Aptos")
    $fontSize = Get-Number (Get-Value $TextSpec "fontSize" 14) 14
    $color = Get-Value $TextSpec "color" "#1E293B"
    $bold = Get-BoolAsMso (Get-Value $TextSpec "bold" $false) $false
    $align = Get-TextAlignment (Get-Value $TextSpec "align" "center")
    $valign = Get-VerticalAnchor (Get-Value $TextSpec "valign" "middle")

    $Shape.TextFrame.TextRange.Text = [string]$text
    $Shape.TextFrame.TextRange.Font.Name = [string]$font
    $Shape.TextFrame.TextRange.Font.Size = $fontSize
    $Shape.TextFrame.TextRange.Font.Color.RGB = ConvertTo-PptColor $color "#1E293B"
    $Shape.TextFrame.TextRange.Font.Bold = $bold
    $Shape.TextFrame.TextRange.ParagraphFormat.Alignment = $align
    $Shape.TextFrame.VerticalAnchor = $valign
    $Shape.TextFrame.WordWrap = $script:MsoTrue
    $Shape.TextFrame.MarginLeft = Get-Number (Get-Value $TextSpec "marginLeft" 6) 6
    $Shape.TextFrame.MarginRight = Get-Number (Get-Value $TextSpec "marginRight" 6) 6
    $Shape.TextFrame.MarginTop = Get-Number (Get-Value $TextSpec "marginTop" 3) 3
    $Shape.TextFrame.MarginBottom = Get-Number (Get-Value $TextSpec "marginBottom" 3) 3
}

function Set-ShapeIdentity {
    param(
        [object]$Shape,
        [object]$Element
    )

    $id = Get-Value $Element "id" $null
    if ($null -ne $id -and [string]$id -ne "") {
        $Shape.Name = [string]$id
        $Shape.AlternativeText = [string]$id
        $script:ShapeById[[string]$id] = $Shape
    }
}

function Add-Box {
    param(
        [object]$Slide,
        [object]$Element,
        [object]$Theme
    )

    $x = Get-Number (Get-Value $Element "x" 0)
    $y = Get-Number (Get-Value $Element "y" 0)
    $w = Get-Number (Get-Value $Element "w" 120)
    $h = Get-Number (Get-Value $Element "h" 48)
    $shapeType = Get-ShapeType (Get-Value $Element "shape" "roundRect")

    $shape = $Slide.Shapes.AddShape($shapeType, $x, $y, $w, $h)
    Set-ShapeIdentity $shape $Element
    Set-FillStyle $shape (Get-Value $Element "fill" "#FFFFFF") "#FFFFFF"
    Set-LineStyle $shape (Get-Value $Element "line" $null) "#C9D3E0"

    $textSpec = Get-Value $Element "text" $null
    if ($textSpec -is [string]) {
        $textSpec = @{
            text = $textSpec
            fontSize = Get-Value $Element "fontSize" 13
            font = Get-Value $Element "font" (Get-Value $Theme "font" "Aptos")
            color = Get-Value $Element "color" "#1E293B"
            bold = Get-Value $Element "bold" $false
            align = Get-Value $Element "align" "center"
            valign = Get-Value $Element "valign" "middle"
        }
    }
    Set-TextStyle $shape $textSpec $Theme
    Wait-DrawStep
}

function Add-Text {
    param(
        [object]$Slide,
        [object]$Element,
        [object]$Theme
    )

    $x = Get-Number (Get-Value $Element "x" 0)
    $y = Get-Number (Get-Value $Element "y" 0)
    $w = Get-Number (Get-Value $Element "w" 160)
    $h = Get-Number (Get-Value $Element "h" 28)
    $shape = $Slide.Shapes.AddTextbox($script:MsoTextOrientationHorizontal, $x, $y, $w, $h)
    Set-ShapeIdentity $shape $Element
    $shape.Fill.Visible = $script:MsoFalse
    $shape.Line.Visible = $script:MsoFalse
    Set-TextStyle $shape $Element $Theme
    Wait-DrawStep
}

function Add-Line {
    param(
        [object]$Slide,
        [object]$Element
    )

    $from = Get-Value $Element "from" $null
    $to = Get-Value $Element "to" $null

    if ($null -ne $from -and $null -ne $to) {
        $fromArray = @($from)
        $toArray = @($to)
        $x1 = Get-Number $fromArray[0]
        $y1 = Get-Number $fromArray[1]
        $x2 = Get-Number $toArray[0]
        $y2 = Get-Number $toArray[1]
    }
    else {
        $x1 = Get-Number (Get-Value $Element "x1" 0)
        $y1 = Get-Number (Get-Value $Element "y1" 0)
        $x2 = Get-Number (Get-Value $Element "x2" 100)
        $y2 = Get-Number (Get-Value $Element "y2" 100)
    }

    $connectorType = Get-ConnectorType (Get-Value $Element "connector" "straight")
    $shape = $Slide.Shapes.AddConnector($connectorType, $x1, $y1, $x2, $y2)
    Set-ShapeIdentity $shape $Element
    Set-LineStyle $shape (Get-Value $Element "line" $Element) "#64748B"
    Wait-DrawStep
}

function Add-Image {
    param(
        [object]$Slide,
        [object]$Element,
        [string]$SpecDir,
        [hashtable]$AssetById
    )

    $path = Resolve-ImagePath $Element $SpecDir $AssetById
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Image not found: $path"
    }

    $x = Get-Number (Get-Value $Element "x" 0)
    $y = Get-Number (Get-Value $Element "y" 0)
    $w = Get-Number (Get-Value $Element "w" 120)
    $h = Get-Number (Get-Value $Element "h" 90)
    $fit = ([string](Get-Value $Element "fit" "cover")).ToLowerInvariant()

    $frame = Get-Value $Element "frame" $null
    if ($null -ne $frame) {
        $frameShape = $Slide.Shapes.AddShape((Get-ShapeType (Get-Value $frame "shape" "rect")), $x, $y, $w, $h)
        Set-FillStyle $frameShape (Get-Value $frame "fill" "none") "none"
        Set-LineStyle $frameShape (Get-Value $frame "line" "none") "#E2E8F0"
    }

    if ($fit -eq "contain") {
        $shape = $Slide.Shapes.AddPicture($path, $script:MsoFalse, $script:MsoTrue, $x, $y, -1, -1)
        $originalW = [double]$shape.Width
        $originalH = [double]$shape.Height
        if ($originalW -le 0 -or $originalH -le 0) {
            throw "Could not read image dimensions: $path"
        }

        $scale = [Math]::Min($w / $originalW, $h / $originalH)
        $shape.Width = $originalW * $scale
        $shape.Height = $originalH * $scale
        $shape.Left = $x + (($w - $shape.Width) / 2)
        $shape.Top = $y + (($h - $shape.Height) / 2)
    }
    elseif ($fit -eq "stretch") {
        $shape = $Slide.Shapes.AddPicture($path, $script:MsoFalse, $script:MsoTrue, $x, $y, $w, $h)
    }
    elseif ($fit -eq "original") {
        $shape = $Slide.Shapes.AddPicture($path, $script:MsoFalse, $script:MsoTrue, $x, $y, -1, -1)
    }
    else {
        $shape = $Slide.Shapes.AddPicture($path, $script:MsoFalse, $script:MsoTrue, $x, $y, -1, -1)
        $originalW = [double]$shape.Width
        $originalH = [double]$shape.Height
        if ($originalW -le 0 -or $originalH -le 0) {
            throw "Could not read image dimensions: $path"
        }

        $frameRatio = $w / $h
        $imageRatio = $originalW / $originalH
        if ($imageRatio -gt $frameRatio) {
            $visibleW = $originalH * $frameRatio
            $crop = ($originalW - $visibleW) / 2
            $shape.PictureFormat.CropLeft = $crop
            $shape.PictureFormat.CropRight = $crop
        }
        elseif ($imageRatio -lt $frameRatio) {
            $visibleH = $originalW / $frameRatio
            $crop = ($originalH - $visibleH) / 2
            $shape.PictureFormat.CropTop = $crop
            $shape.PictureFormat.CropBottom = $crop
        }
        $shape.Left = $x
        $shape.Top = $y
        $shape.Width = $w
        $shape.Height = $h
    }

    Set-ShapeIdentity $shape $Element
    $line = Get-Value $Element "line" $null
    if ($null -ne $line) {
        Set-LineStyle $shape $line "#E2E8F0"
    }
    Wait-DrawStep
}

function Add-Element {
    param(
        [object]$Slide,
        [object]$Element,
        [object]$Theme,
        [string]$SpecDir,
        [hashtable]$AssetById
    )

    $type = ([string](Get-Value $Element "type" "box")).ToLowerInvariant()
    $id = Get-Value $Element "id" ""
    $label = Get-Value $Element "text" ""
    if ($label -isnot [string]) {
        $label = Get-Value $label "text" ""
    }

    try {
        switch ($type) {
            "box" { Add-Box $Slide $Element $Theme }
            "shape" { Add-Box $Slide $Element $Theme }
            "text" { Add-Text $Slide $Element $Theme }
            "title" {
                if ($null -eq (Get-Value $Element "bold" $null)) {
                    $Element | Add-Member -NotePropertyName "bold" -NotePropertyValue $true -Force
                }
                if ($null -eq (Get-Value $Element "fontSize" $null)) {
                    $Element | Add-Member -NotePropertyName "fontSize" -NotePropertyValue 20 -Force
                }
                Add-Text $Slide $Element $Theme
            }
            "line" { Add-Line $Slide $Element }
            "connector" { Add-Line $Slide $Element }
            "image" { Add-Image $Slide $Element $SpecDir $AssetById }
            "sleep" {
                $duration = [int](Get-Number (Get-Value $Element "ms" $script:StepDelayMs) $script:StepDelayMs)
                if ($duration -gt 0) {
                    Start-Sleep -Milliseconds $duration
                }
            }
            default { throw "Unsupported element type '$type'." }
        }
    }
    catch {
        throw "Failed element type='$type' id='$id' text='$label': $($_.Exception.Message)"
    }
}

function Test-DeckSpec {
    param(
        [object]$DeckSpec,
        [string]$SpecDir,
        [hashtable]$AssetById
    )

    $slides = @(Get-Value $DeckSpec "slides" @())
    if ($slides.Count -eq 0) {
        throw "Spec contains no slides."
    }

    foreach ($assetId in $AssetById.Keys) {
        $asset = $AssetById[$assetId]
        $path = Get-Value $asset "resolvedPath" $null
        if ($null -ne $path -and [string]$path -ne "" -and -not (Test-Path -LiteralPath $path)) {
            throw "Asset '$assetId' points to a missing file: $path"
        }
    }

    $slideNumber = 0
    foreach ($slideSpec in $slides) {
        $slideNumber += 1
        $elementNumber = 0
        foreach ($element in @(Get-Value $slideSpec "elements" @())) {
            $elementNumber += 1
            $type = ([string](Get-Value $element "type" "box")).ToLowerInvariant()
            if ($type -eq "image") {
                try {
                    $path = Resolve-ImagePath $element $SpecDir $AssetById
                    if (-not (Test-Path -LiteralPath $path)) {
                        throw "Image not found: $path"
                    }
                }
                catch {
                    throw "Slide $slideNumber element $elementNumber image validation failed: $($_.Exception.Message)"
                }
            }
        }
    }

    return $true
}

$resolvedSpec = Resolve-Path -LiteralPath $Spec
$specDir = Split-Path -Parent $resolvedSpec
$deckSpec = Get-Content -Raw -LiteralPath $resolvedSpec | ConvertFrom-Json
$assetById = Build-AssetMap $deckSpec $specDir
Test-DeckSpec $deckSpec $specDir $assetById | Out-Null

if ($ValidateOnly) {
    Write-Host "Spec validation OK."
    Write-Host "Slides: $(@(Get-Value $deckSpec "slides" @()).Count)"
    Write-Host "Assets: $($assetById.Count)"
    exit 0
}

$page = Get-Value $deckSpec "page" $null
$theme = Get-Value $deckSpec "theme" $null
if ($null -eq $theme) {
    $theme = Get-Value $page "theme" @{ font = "Aptos" }
}

if ([string]::IsNullOrWhiteSpace($Output)) {
    $Output = Get-Value $deckSpec "output" "exports/live-office-ppt.pptx"
}
$outputPath = Resolve-ElementPath $Output $specDir
$outputDir = Split-Path -Parent $outputPath
if (-not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
}
if ((Test-Path -LiteralPath $outputPath) -and -not $Overwrite) {
    throw "Output exists: $outputPath. Pass -Overwrite to replace it."
}

Write-Host "Opening PowerPoint..."
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = if ($Visible) { $MsoTrue } else { $MsoFalse }
$presentation = $ppt.Presentations.Add($MsoTrue)

$width = Get-Number (Get-Value $page "width" 960) 960
$height = Get-Number (Get-Value $page "height" 540) 540
$presentation.PageSetup.SlideWidth = $width
$presentation.PageSetup.SlideHeight = $height

$slides = @(Get-Value $deckSpec "slides" @())
if ($slides.Count -eq 0) {
    throw "Spec contains no slides."
}

$slideIndex = 0
foreach ($slideSpec in $slides) {
    $slideIndex += 1
    $script:ShapeById = @{}
    Write-Host "Drawing slide $slideIndex / $($slides.Count)..."
    $slide = $presentation.Slides.Add($slideIndex, $PpLayoutBlank)

    $background = Get-Value $slideSpec "background" (Get-Value $page "background" "#FFFFFF")
    $slide.FollowMasterBackground = $MsoFalse
    $slide.Background.Fill.ForeColor.RGB = ConvertTo-PptColor $background "#FFFFFF"

    foreach ($element in @(Get-Value $slideSpec "elements" @())) {
        Add-Element $slide $element $theme $specDir $assetById
    }

    $notes = Get-Value $slideSpec "notes" $null
    if ($null -ne $notes -and [string]$notes -ne "") {
        try {
            $slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = [string]$notes
        }
        catch {
            Write-Warning "Could not write notes on slide ${slideIndex}: $($_.Exception.Message)"
        }
    }
}

Write-Host "Saving deck: $outputPath"
$presentation.SaveAs($outputPath, $PpSaveAsOpenXmlPresentation)

if ((-not $KeepOpen) -or $CloseAfterSave) {
    $presentation.Close()
    $ppt.Quit()
}

Write-Host "Done."
Write-Host $outputPath
