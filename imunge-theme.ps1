<# 
.SYNOPSIS
  Interactively replace CT_RAW colors in a Visual Studio .vstheme file.

.DESCRIPTION
  - Edits the .vstheme IN PLACE (no OutputFile).
  - Prompts for SourceColor values in a loop until you enter a blank line.
  - For each SourceColor, TargetColor := SourceColor; picks the closest unused RGB.
  - Preserves original alpha (AA in AARRGGBB) and textual formatting/entities.
  - Appends CSV lines to -LogFile: "category,ui name,original color,new color".

.PARAMETER InputFile
  Path to the .vstheme XML file to modify (in place).

.PARAMETER LogFile
  Fully-qualified path to the log file to append (CSV lines).

.EXAMPLE
  .\Replace-VsThemeColor-Interactive.ps1 -InputFile C:\themes\My.vstheme -LogFile C:\logs\changes.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputFile,

    [Parameter(Mandatory = $true)]
    [string]$LogFile
)

#region Utilities

function Get-FullPathStrict
{
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ParamName
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { throw "$ParamName cannot be empty." }
    $isFullyQualified = $Path -match '^(?:[A-Za-z]:[\\/]|\\\\|//)'
    if (-not $isFullyQualified)
    {
        throw "$ParamName must be a fully qualified path (e.g., C:\folder\file.ext or \\server\share\file.ext)."
    }
    [System.IO.Path]::GetFullPath($Path)
}

function Ensure-ParentDirectory
{
    param([string]$FilePath)
    $full = [System.IO.Path]::GetFullPath($FilePath)
    $parent = [System.IO.Path]::GetDirectoryName($full)
    if ($parent -and -not (Test-Path -LiteralPath $parent))
    {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $full
}

function Test-HexRgb
{
    param([string]$Rgb)
    return ($Rgb -match '^#?[0-9A-Fa-f]{6}$')
}
function Get-RgbBytesFromHash
{
    param([string]$Rgb) # "#RRGGBB" or "RRGGBB"
    if (-not (Test-HexRgb $Rgb))
    {
        throw "Invalid RGB color '$Rgb'. Expected #RRGGBB (leading '#' optional)."
    }
    if ($Rgb.StartsWith('#')) { $Rgb = $Rgb.Substring(1) }
    
    ([Convert]::ToByte($Rgb.Substring(0, 2), 16)),
    ([Convert]::ToByte($Rgb.Substring(2, 2), 16)),
    ([Convert]::ToByte($Rgb.Substring(4, 2), 16))
}

function Split-ArgbHex8
{
    param([string]$ArgbHex8) # "AARRGGBB"
    if ($ArgbHex8 -notmatch '^[0-9A-Fa-f]{8}$') { throw "Invalid ARGB '$ArgbHex8'." }
    $a = [Convert]::ToByte($ArgbHex8.Substring(0, 2), 16)
    $r = [Convert]::ToByte($ArgbHex8.Substring(2, 2), 16)
    $g = [Convert]::ToByte($ArgbHex8.Substring(4, 2), 16)
    $b = [Convert]::ToByte($ArgbHex8.Substring(6, 2), 16)
    [byte[]]@($a, $r, $g, $b)
}
function Join-ArgbHex8 { param([byte]$A, [byte]$R, [byte]$G, [byte]$B) ('{0:X2}{1:X2}{2:X2}{3:X2}' -f $A, $R, $G, $B) }

function Clamp-Byte { param([int]$x) if ($x -lt 0) { [byte]0 } elseif ($x -gt 255) { [byte]255 } else { [byte]$x } }
function Get-EuclideanSq
{
    param([byte]$R1, [byte]$G1, [byte]$B1, [byte]$R2, [byte]$G2, [byte]$B2)
    $dr = [int]$R1 - [int]$R2; $dg = [int]$G1 - [int]$G2; $db = [int]$B1 - [int]$B2
    $dr * $dr + $dg * $dg + $db * $db
}
function Find-ClosestUnusedColor
{
    param([byte]$TR, [byte]$TG, [byte]$TB, [System.Collections.Generic.HashSet[string]]$UsedRgbSet)
    $targetKey = ('{0:X2}{1:X2}{2:X2}' -f $TR, $TG, $TB).ToLowerInvariant()
    if (-not $UsedRgbSet.Contains($targetKey)) { return , $TR, $TG, $TB }

    $dirs = @(); foreach ($dx in -1..1)
    {
        foreach ($dy in -1..1)
        {
            foreach ($dz in -1..1)
            {
                if ($dx -eq 0 -and $dy -eq 0 -and $dz -eq 0) { continue }; $dirs += , @($dx, $dy, $dz)
            }
        }
    } 
    $best = $null; $bestDist = [int]::MaxValue
    for ($rad = 1; $rad -le 255; $rad++)
    {
        foreach ($d in $dirs)
        {
            $r = Clamp-Byte([int]$TR + $rad * $d[0]); $g = Clamp-Byte([int]$TG + $rad * $d[1]); $b = Clamp-Byte([int]$TB + $rad * $d[2])
            $key = ('{0:X2}{1:X2}{2:X2}' -f $r, $g, $b).ToLowerInvariant()
            if ($UsedRgbSet.Contains($key)) { continue }
            $dist = Get-EuclideanSq -R1 $r -G1 $g -B1 $b -R2 $TR -G2 $TG -B2 $TB
            if ($dist -lt $bestDist) { $bestDist = $dist; $best = @([byte]$r, [byte]$g, [byte]$b) }
        }
        if ($best) { break }
    }
    if (-not $best) { throw "Failed to find an unused color." }
    $best
}

# Regex for <Background>/<Foreground> start tag with Type="CT_RAW" and Source="AARRGGBB"
$TagWithRawAndSourcePattern = '(?is)<(?<tag>Background|Foreground)\b(?=[^>]*\bType\s*=\s*["'']CT_RAW["''])(?=[^>]*\bSource\s*=\s*["''][0-9A-Fa-f]{8}["''])[^>]*?\bSource\s*=\s*(?<pre>["''])(?<hex>[0-9A-Fa-f]{8})(?<post>["''])'

function Get-UsedRgbSetFromText
{
    param([string]$XmlText)
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    $re = [System.Text.RegularExpressions.Regex]::new($TagWithRawAndSourcePattern, 'IgnoreCase,Singleline')
    foreach ($m in $re.Matches($XmlText))
    {
        $hex = $m.Groups['hex'].Value; $rgbKey = $hex.Substring(2, 6).ToLowerInvariant(); [void]$set.Add($rgbKey)
    }
    $set
}

# Capture <Color Name="...">...</Color> spans to attribute UI names
function Get-ColorSpansFromText
{
    param([string]$XmlText)
    $re = [System.Text.RegularExpressions.Regex]::new('(?is)<Color\b(?<attrs>[^>]*)>(?<content>.*?)</Color>')
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($m in $re.Matches($XmlText))
    {
        $attrs = $m.Groups['attrs'].Value; $name = '(unnamed)'
        $nm = [regex]::Match($attrs, '(?is)\bName\s*=\s*(["''])(?<n>.*?)\1'); if ($nm.Success) { $name = $nm.Groups['n'].Value }
        $list.Add([pscustomobject]@{Start = $m.Index; End = $m.Index + $m.Length; Name = $name }) | Out-Null
    }
    $list
}
function Get-UiElementNameByIndex
{
    param([int]$Index, [System.Collections.Generic.List[object]]$ColorSpans)
    foreach ($s in $ColorSpans) { if ($Index -ge $s.Start -and $Index -lt $s.End) { return $s.Name } }
    '(unnamed)'
}

# Capture <Category Name="...">...</Category> spans to attribute Category names
function Get-CategorySpansFromText
{
    param([string]$XmlText)
    $re = [System.Text.RegularExpressions.Regex]::new('(?is)<Category\b(?<attrs>[^>]*)>(?<content>.*?)</Category>')
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($m in $re.Matches($XmlText))
    {
        $attrs = $m.Groups['attrs'].Value; $name = '(uncategorized)'
        $nm = [regex]::Match($attrs, '(?is)\bName\s*=\s*(["''])(?<n>.*?)\1'); if ($nm.Success) { $name = $nm.Groups['n'].Value }
        $list.Add([pscustomobject]@{Start = $m.Index; End = $m.Index + $m.Length; Name = $name }) | Out-Null
    }
    $list
}
function Get-CategoryNameByIndex
{
    param([int]$Index, [System.Collections.Generic.List[object]]$CategorySpans)
    foreach ($s in $CategorySpans) { if ($Index -ge $s.Start -and $Index -lt $s.End) { return $s.Name } }
    '(uncategorized)'
}

function Replace-MatchingSourcesInText
{
    param(
        [string]$XmlText,
        [byte[]]$SourceRgb,  # [R,G,B]
        [byte[]]$TargetRgb,  # [R,G,B]  (here: same as SourceRgb)
        [System.Collections.Generic.HashSet[string]]$UsedRgbSet,
        [ref]$ReplacedCount,
        [System.Collections.Generic.List[string]]$LogBuffer,
        [System.Collections.Generic.List[object]]$ColorSpans,
        [System.Collections.Generic.List[object]]$CategorySpans
    )

    $re = [System.Text.RegularExpressions.Regex]::new($TagWithRawAndSourcePattern, 'IgnoreCase,Singleline')

    # Capture needed values locally for the delegate (no $using:)
    $srcKey = ('{0:X2}{1:X2}{2:X2}' -f $SourceRgb[0], $SourceRgb[1], $SourceRgb[2]).ToLowerInvariant()
    $tR = [byte]$TargetRgb[0]
    $tG = [byte]$TargetRgb[1]
    $tB = [byte]$TargetRgb[2]
    $usedSet = $UsedRgbSet
    $logBuf = $LogBuffer
    $cSpans = $ColorSpans
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

        # Find closest unused to Target (=Source)
        $best = Find-ClosestUnusedColor -TR $tR -TG $tG -TB $tB -UsedRgbSet $usedSet
        $newR = [byte]$best[0]; $newG = [byte]$best[1]; $newB = [byte]$best[2]
        $newKey = ('{0:X2}{1:X2}{2:X2}' -f $newR, $newG, $newB).ToLowerInvariant()
        [void]$usedSet.Add($newKey)

        $newHex = Join-ArgbHex8 -A $a -R $newR -G $newG -B $newB

        # Names for logging
        $uiName = Get-UiElementNameByIndex -Index $m.Index -ColorSpans $cSpans
        $catName = Get-CategoryNameByIndex -Index $m.Index -CategorySpans $catSpans
        $origRgbHash = ('#{0:X2}{1:X2}{2:X2}' -f $r, $g, $b)
        $newRgbHash = ('#{0:X2}{1:X2}{2:X2}' -f $newR, $newG, $newB)

        # CSV: category,ui name,original color,new color
        $logBuf.Add(("{0},{1},{2},{3}" -f $catName, $uiName, $origRgbHash, $newRgbHash)) | Out-Null

        $script:__replcount = $script:__replcount + 1

        # Replace only the 8 hex digits
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
    if (-not (Test-Path -LiteralPath $InputFile)) { throw "Input file not found: $InputFile" }

    # Enforce LogFile full path, and normalize paths
    $LogFile = Get-FullPathStrict -Path $LogFile   -ParamName 'LogFile'
    $InputFile = [System.IO.Path]::GetFullPath($InputFile)

    # Load theme text (preserve entities)
    $text = Get-Content -LiteralPath $InputFile -Raw -ErrorAction Stop

    Write-Host "Interactive mode. Enter Source colors to replace; press Enter on a blank line to finish."
    $totalReplaced = 0
    $logToAppend = New-Object System.Collections.Generic.List[string]

    while ($true)
    {
        $srcInput = Read-Host "Source color (#RRGGBB)"
        if ([string]::IsNullOrWhiteSpace($srcInput)) { break }

        # Parse color and set Target = Source
        $srcRgb = Get-RgbBytesFromHash $srcInput
        $tgtRgb = $srcRgb

        # Build used-color set & spans from CURRENT text (so replacements remain unique)
        $used = Get-UsedRgbSetFromText -XmlText $text
        $colorSpans = Get-ColorSpansFromText    -XmlText $text
        $categorySpans = Get-CategorySpansFromText -XmlText $text

        $countRef = 0
        $sessionLog = New-Object System.Collections.Generic.List[string]

        $text = Replace-MatchingSourcesInText `
            -XmlText $text -SourceRgb $srcRgb -TargetRgb $tgtRgb -UsedRgbSet $used `
            -ReplacedCount ([ref]$countRef) -LogBuffer $sessionLog `
            -ColorSpans $colorSpans -CategorySpans $categorySpans

        $totalReplaced += $countRef
        if ($sessionLog.Count -gt 0)
        {
            $logToAppend.AddRange($sessionLog)
        }
        Write-Host ("Replaced {0} occurrence(s) for {1}." -f $countRef, ($srcInput))
    }

    # Write updated theme back IN PLACE
    Set-Content -LiteralPath $InputFile -Value $text -Encoding UTF8

    # Append log lines (once)
    if ($logToAppend.Count -gt 0)
    {
        $logPath = Ensure-ParentDirectory -FilePath $LogFile
        Add-Content -LiteralPath $logPath -Value $logToAppend -Encoding UTF8
        Write-Host ("Done. Total replacements: {0}`nLog: {1}" -f $totalReplaced, $logPath)
    }
    else
    {
        Write-Host ("Done. No changes made.")
    }
}
catch
{
    Write-Error $_.Exception.Message
    exit 1
}

#endregion
