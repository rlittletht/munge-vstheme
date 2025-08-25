<# 
.SYNOPSIS
  Replace specific CT_RAW colors in a Visual Studio .vstheme XML file.

.DESCRIPTION
  - Reads an input .vstheme (XML).
  - Collects the set of "in-use" colors from all Background/Foreground elements with Type="CT_RAW".
  - For each CT_RAW whose RGB equals -SourceColor (#RRGGBB), replaces its RGB with a visually close color based on -TargetColor,
    ensuring the chosen RGB is NOT already used. The original alpha byte in each element is preserved.
  - Writes the modified XML to -OutputFile.

.PARAMETER InputFile
  Path to the source .vstheme XML file.

.PARAMETER OutputFile
  Path to write the modified .vstheme XML.

.PARAMETER SourceColor
  Color to find, in form "#RRGGBB".

.PARAMETER TargetColor
  Desired color basis, in form "#RRGGBB". If this exact RGB is already used, the script finds the closest unused RGB.

.PARAMETER LogFile
 File to record changes into

.EXAMPLE
  .\Replace-VsThemeColor.ps1 -InputFile .\MyTheme.vstheme -OutputFile .\MyTheme-Updated.vstheme -SourceColor "#31FF00" -TargetColor "#30EE10"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputFile,

    [Parameter(Mandatory = $true)]
    [string]$OutputFile,

    [Parameter(Mandatory = $true)]
    [string]$SourceColor,

    [Parameter(Mandatory = $false)]
    [string]$TargetColor = $null,
    
    [Parameter(Mandatory = $true)]
    [string]$LogFile        # path to append change logs
)

#region Utility: Parsing & Validation

function Test-HexRgb
{
    param([string]$Rgb)
    # Accept with or without leading '#'
    return ($Rgb -match '^#?[0-9A-Fa-f]{6}$')
}

function Get-RgbBytesFromHash
{
    param([string]$Rgb) # "#RRGGBB" or "RRGGBB"
    if (-not (Test-HexRgb $Rgb))
    {
        throw "Invalid RGB color '$Rgb'. Expected format: #RRGGBB (leading '#' optional)."
    }
    if ($Rgb.StartsWith('#')) { $Rgb = $Rgb.Substring(1) }
    $rr = [Convert]::ToByte($Rgb.Substring(0, 2), 16)
    $gg = [Convert]::ToByte($Rgb.Substring(2, 2), 16)
    $bb = [Convert]::ToByte($Rgb.Substring(4, 2), 16)
    [byte[]]@($rr, $gg, $bb)
}

function Get-HashFromRgbBytes
{
    param([byte]$R, [byte]$G, [byte]$B)
    return ('#{0:X2}{1:X2}{2:X2}' -f $R, $G, $B)
}

function Split-ArgbHex8
{
    param([string]$ArgbHex8) # "AARRGGBB" (no '#')
    if ($ArgbHex8 -notmatch '^[0-9A-Fa-f]{8}$')
    {
        throw "Invalid ARGB '$ArgbHex8'. Expected 8 hex chars (AARRGGBB)."
    }
    $a = [Convert]::ToByte($ArgbHex8.Substring(0, 2), 16)
    $r = [Convert]::ToByte($ArgbHex8.Substring(2, 2), 16)
    $g = [Convert]::ToByte($ArgbHex8.Substring(4, 2), 16)
    $b = [Convert]::ToByte($ArgbHex8.Substring(6, 2), 16)
    [byte[]]@($a, $r, $g, $b)
}

function Join-ArgbHex8
{
    param([byte]$A, [byte]$R, [byte]$G, [byte]$B)
    return ('{0:X2}{1:X2}{2:X2}{3:X2}' -f $A, $R, $G, $B)
}

function Ensure-ParentDirectory
{
    param([string]$FilePath)
    if ([string]::IsNullOrWhiteSpace($FilePath)) { throw "Path cannot be empty." }
    $full = [System.IO.Path]::GetFullPath($FilePath)
    $parent = [System.IO.Path]::GetDirectoryName($full)
    if ($parent -and -not (Test-Path -LiteralPath $parent))
    {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $full
}

function Get-FullPathStrict
{
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ParamName
    )
    if ([string]::IsNullOrWhiteSpace($Path))
    {
        throw "$ParamName cannot be empty."
    }
    # Must start with drive-root (C:\...) or UNC (\\server\share\...)
    $isFullyQualified = $Path -match '^(?:[A-Za-z]:[\\/]|\\\\|//)'
    if (-not $isFullyQualified)
    {
        throw "$ParamName must be a fully qualified path (e.g., C:\folder\file.ext or \\server\share\file.ext)."
    }
    return [System.IO.Path]::GetFullPath($Path)
}

#endregion

#region Color math & selection

function Get-EuclideanSq
{
    param(
        [byte]$R1, [byte]$G1, [byte]$B1,
        [byte]$R2, [byte]$G2, [byte]$B2
    )
    # squared distance in RGB space
    $dr = [int]$R1 - [int]$R2
    $dg = [int]$G1 - [int]$G2
    $db = [int]$B1 - [int]$B2
    return ($dr * $dr + $dg * $dg + $db * $db)
}

function Clamp-Byte
{
    param([int]$x)
    if ($x -lt 0) { return [byte]0 }
    if ($x -gt 255) { return [byte]255 }
    return [byte]$x
}

# Generates up to 26*radius candidates per radius step by stepping along axis and diagonal unit directions.
# This gives a good approximation for "closest" in Euclidean space without exploring all 16M colors.
function Find-ClosestUnusedColor
{
    param(
        [byte]$TR, [byte]$TG, [byte]$TB,
        [System.Collections.Generic.HashSet[string]]$UsedRgbSet # contains "rrggbb" lower-case
    )

    $targetKey = ('{0:X2}{1:X2}{2:X2}' -f $TR, $TG, $TB).ToLowerInvariant()
    if (-not $UsedRgbSet.Contains($targetKey))
    {
        # Target is free — use it
        return , $TR, $TG, $TB
    }

    # Directions = all non-zero combinations of (-1,0,1) for R,G,B (26 directions)
    $dirs = @()
    foreach ($dx in -1..1)
    {
        foreach ($dy in -1..1)
        {
            foreach ($dz in -1..1)
            {
                if ($dx -eq 0 -and $dy -eq 0 -and $dz -eq 0) { continue }
                $dirs += , @($dx, $dy, $dz)
            }
        }
    }

    $best = $null
    $bestDist = [int]::MaxValue

    for ($rad = 1; $rad -le 255; $rad++)
    {
        foreach ($d in $dirs)
        {
            $rCand = Clamp-Byte([int]$TR + $rad * $d[0])
            $gCand = Clamp-Byte([int]$TG + $rad * $d[1])
            $bCand = Clamp-Byte([int]$TB + $rad * $d[2])

            $key = ('{0:X2}{1:X2}{2:X2}' -f $rCand, $gCand, $bCand).ToLowerInvariant()
            if ($UsedRgbSet.Contains($key)) { continue }

            $dist = Get-EuclideanSq -R1 $rCand -G1 $gCand -B1 $bCand -R2 $TR -G2 $TG -B2 $TB
            if ($dist -lt $bestDist)
            {
                $bestDist = $dist
                $best = @([byte]$rCand, [byte]$gCand, [byte]$bCand)
                if ($bestDist -eq 0) { break } # exact (shouldn't happen, already checked), but just in case
            }
        }
        if ($best) { break }
    }

    if (-not $best)
    {
        throw "Failed to find an unused color (this is extremely unlikely)."
    }

    return $best
}

#endregion


#region Text parsing (regex-only, preserves entities)

# Pattern:
#  - Matches <Background> or <Foreground> start tag
#  - Asserts it contains Type="CT_RAW" AND Source="AARRGGBB"
#  - Captures:
#     * tag  : Background|Foreground
#     * pre  : opening quote of Source
#     * hex  : 8 hex digits (AARRGGBB)
#     * post : closing quote
# Notes:
#  - order-insensitive for attributes
#  - supports single or double quotes
$TagWithRawAndSourcePattern = '(?is)<(?<tag>Background|Foreground)\b(?=[^>]*\bType\s*=\s*["'']CT_RAW["''])(?=[^>]*\bSource\s*=\s*["''][0-9A-Fa-f]{8}["''])[^>]*?\bSource\s*=\s*(?<pre>["''])(?<hex>[0-9A-Fa-f]{8})(?<post>["''])'

function Get-UsedRgbSetFromText
{
    param([string]$XmlText)

    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    $regex = [System.Text.RegularExpressions.Regex]::new($TagWithRawAndSourcePattern, 'IgnoreCase, Singleline')

    foreach ($m in $regex.Matches($XmlText))
    {
        $hex = $m.Groups['hex'].Value  # AARRGGBB
        $rgbKey = $hex.Substring(2, 6).ToLowerInvariant()
        [void]$set.Add($rgbKey)
    }
    $set
}

function Get-CategorySpansFromText
{
    param([string]$XmlText)

    # Match each <Category ...>...</Category> block (non-greedy)
    $re = [System.Text.RegularExpressions.Regex]::new('(?is)<Category\b(?<attrs>[^>]*)>(?<content>.*?)</Category>')
    $spans = New-Object System.Collections.Generic.List[object]

    foreach ($m in $re.Matches($XmlText))
    {
        $attrs = $m.Groups['attrs'].Value
        $name = '(uncategorized)'
        $nm = [regex]::Match($attrs, '(?is)\bName\s*=\s*(["''])(?<n>.*?)\1')
        if ($nm.Success) { $name = $nm.Groups['n'].Value }

        $spans.Add([pscustomobject]@{
                Start = $m.Index
                End   = $m.Index + $m.Length  # exclusive
                Name  = $name
            }) | Out-Null
    }
    $spans
}

function Get-CategoryNameByIndex
{
    param(
        [int]$Index,
        [System.Collections.Generic.List[object]]$CategorySpans
    )
    foreach ($s in $CategorySpans)
    {
        if ($Index -ge $s.Start -and $Index -lt $s.End) { return $s.Name }
    }
    return '(uncategorized)'
}

function Get-ColorSpansFromText
{
    param([string]$XmlText)

    # Match each <Color ...>...</Color> block (non-greedy), regardless of attribute order
    $re = [System.Text.RegularExpressions.Regex]::new('(?is)<Color\b(?<attrs>[^>]*)>(?<content>.*?)</Color>')
    $spans = New-Object System.Collections.Generic.List[object]

    foreach ($m in $re.Matches($XmlText))
    {
        $attrs = $m.Groups['attrs'].Value
        $name = '(unnamed)'
        $nm = [regex]::Match($attrs, '(?is)\bName\s*=\s*(["''])(?<n>.*?)\1')
        if ($nm.Success) { $name = $nm.Groups['n'].Value }

        $spans.Add([pscustomobject]@{
                Start = $m.Index
                End   = $m.Index + $m.Length  # exclusive
                Name  = $name
            }) | Out-Null
    }
    $spans
}

function Get-UiElementNameByIndex
{
    param(
        [int]$Index,
        [System.Collections.Generic.List[object]]$ColorSpans
    )
    foreach ($s in $ColorSpans)
    {
        if ($Index -ge $s.Start -and $Index -lt $s.End) { return $s.Name }
    }
    return '(unnamed)'
}

function Replace-MatchingSourcesInText
{
    param(
        [string]$XmlText,
        [byte[]]$SourceRgb,  # [R,G,B]
        [byte[]]$TargetRgb,  # [R,G,B]
        [System.Collections.Generic.HashSet[string]]$UsedRgbSet,
        [ref]$ReplacedCount,
        [System.Collections.Generic.List[string]]$LogBuffer,
        [System.Collections.Generic.List[object]]$ColorSpans,
        [System.Collections.Generic.List[object]]$CategorySpans

    )

    $re = [System.Text.RegularExpressions.Regex]::new($TagWithRawAndSourcePattern, 'IgnoreCase, Singleline')

    # Capture needed values locally for the delegate (no $using:)
    $srcKey = ('{0:X2}{1:X2}{2:X2}' -f $SourceRgb[0], $SourceRgb[1], $SourceRgb[2]).ToLowerInvariant()
    $tR = [byte]$TargetRgb[0]
    $tG = [byte]$TargetRgb[1]
    $tB = [byte]$TargetRgb[2]
    $usedSet = $UsedRgbSet
    $logBuf = $LogBuffer
    $spans = $ColorSpans
    $catSpans = $CategorySpans
    
    $evaluator = [System.Text.RegularExpressions.MatchEvaluator] {
        param([System.Text.RegularExpressions.Match]$m)

        $aarrggbb = $m.Groups['hex'].Value
        $a = [Convert]::ToByte($aarrggbb.Substring(0, 2), 16)
        $r = [Convert]::ToByte($aarrggbb.Substring(2, 2), 16)
        $g = [Convert]::ToByte($aarrggbb.Substring(4, 2), 16)
        $b = [Convert]::ToByte($aarrggbb.Substring(6, 2), 16)

        $rgbKeyHere = ('{0:X2}{1:X2}{2:X2}' -f $r, $g, $b).ToLowerInvariant()
        if ($rgbKeyHere -ne $srcKey)
        {
            return $m.Value
        }

        # Choose closest unused to target, then reserve it
        $best = Find-ClosestUnusedColor -TR $tR -TG $tG -TB $tB -UsedRgbSet $usedSet
        $newR = [byte]$best[0]; $newG = [byte]$best[1]; $newB = [byte]$best[2]
        $newKey = ('{0:X2}{1:X2}{2:X2}' -f $newR, $newG, $newB).ToLowerInvariant()
        [void]$usedSet.Add($newKey)

        $newHex = Join-ArgbHex8 -A $a -R $newR -G $newG -B $newB

        # Look up parent <Color Name="..."> by this match's original offset
        $uiName = Get-UiElementNameByIndex -Index $m.Index -ColorSpans $spans
        $catName = Get-CategoryNameByIndex -Index $m.Index -CategorySpans $catSpans
        $tagName = $m.Groups['tag'].Value  # Background/Foreground

        # Log line: category,ui name,original color,new color
        $origRgbHash = ('#{0:X2}{1:X2}{2:X2}' -f $r, $g, $b)
        $newRgbHash = ('#{0:X2}{1:X2}{2:X2}' -f $newR, $newG, $newB)
        $logLine = "{0},{1},{2},{3}" -f $catName, $uiName, $origRgbHash, $newFRgbHash
        $logBuf.Add($logLine) | Out-Null

        $script:__replcount = $script:__replcount + 1

        # Replace only the 8 hex digits inside Source="AARRGGBB"
        return ($m.Value.Remove($m.Groups['hex'].Index - $m.Index, $m.Groups['hex'].Length).Insert($m.Groups['hex'].Index - $m.Index, $newHex))
    }

    $script:__replcount = 0
    $newText = $re.Replace($XmlText, $evaluator)
    $ReplacedCount.Value = $script:__replcount
    return $newText
}


#endregion

#region Main

try
{
    if (-not (Test-Path -LiteralPath $InputFile))
    {
        throw "Input file not found: $InputFile"
    }

    # If -TargetColor wasn’t provided (or is empty), use -SourceColor
    if (-not $PSBoundParameters.ContainsKey('TargetColor') -or [string]::IsNullOrWhiteSpace($TargetColor))
    {
        $TargetColor = $SourceColor
    }

    $OutputFile = Get-FullPathStrict -Path $OutputFile -ParamName 'OutputFile'
    $LogFile = Get-FullPathStrict -Path $LogFile    -ParamName 'LogFile'

    $srcRgb = Get-RgbBytesFromHash $SourceColor
    $tgtRgb = Get-RgbBytesFromHash $TargetColor

    #    [xml]$xml = Get-Content -LiteralPath $InputFile -Raw -ErrorAction Stop
    $text = Get-Content -LiteralPath $InputFile -Raw -ErrorAction Stop
    $colorSpans = Get-ColorSpansFromText -XmlText $text
    $categorySpans = Get-CategorySpansFromText -XmlText $text

    # Collect relevant elements
    #    $elements = [System.Collections.ArrayList](Get-ThemeElements -XmlDoc $xml)

    # Build initial used set from CT_RAW nodes (original colors only)
    #    $used = Build-UsedColorSet -Elements $elements
    $used = Get-UsedRgbSetFromText -XmlText $text

    # Replace matching colors; as we replace, add the new RGBs to the used set to ensure uniqueness
    $countRef = 0
    $logBuffer = New-Object System.Collections.Generic.List[string]

    $updated = Replace-MatchingSourcesInText `
        -XmlText $text `
        -SourceRgb $srcRgb `
        -TargetRgb $tgtRgb `
        -UsedRgbSet $used `
        -ReplacedCount ([ref]$countRef) `
        -LogBuffer $logBuffer `
        -ColorSpans $colorSpans `
        -CategorySpans $categorySpans

    #    $count = Replace-MatchingColors -Elements $elements -SourceRgb $srcRgb -TargetRgb $tgtRgb -UsedRgbSet $used

    $outPath = Ensure-ParentDirectory -FilePath $OutputFile

    #    $xml.Save($OutputFile)
    # Write back, preserving encoding if possible (default is UTF8 without BOM; adjust if you need BOM)
    #    $outPath = Ensure-ParentDirectory -FilePath $OutputFile
    Set-Content -LiteralPath $outPath -Value $updated -Encoding UTF8

    # Write / append logs
    $logPath = Ensure-ParentDirectory -FilePath $LogFile
    if ($logBuffer.Count -gt 0)
    {
        # append lines (create file if it doesn't exist)
        Add-Content -LiteralPath $logPath -Value $logBuffer -Encoding UTF8
    }

    Write-Host ("Done. Replaced {0} occurrence(s) of {1} with colors based on {2}. Output: {3}" -f $countRef, $SourceColor, $TargetColor, $OutputFile)
}
catch
{
    Write-Error $_.Exception.Message
    exit 1
}

#endregion
