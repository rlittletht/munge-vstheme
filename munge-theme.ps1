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

.EXAMPLE
  .\Replace-VsThemeColor.ps1 -InputFile .\MyTheme.vstheme -OutputFile .\MyTheme-Updated.vstheme -SourceColor "#31FF00" -TargetColor "#30EE10"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile,

    [Parameter(Mandatory=$true)]
    [string]$OutputFile,

    [Parameter(Mandatory=$true)]
    [string]$SourceColor,

    [Parameter(Mandatory=$true)]
    [string]$TargetColor
)

#region Utility: Parsing & Validation

function Test-HexRgb {
    param([string]$HashRgb)
    return ($HashRgb -match '^[0-9A-Fa-f]{6}$')
}

function Get-RgbBytesFromHash {
    param([string]$HashRgb) # "RRGGBB"
    if (-not (Test-HexRgb $HashRgb)) {
        throw "Invalid RGB color '$HashRgb'. Expected format: #RRGGBB"
    }
    $rr = [Convert]::ToByte($HashRgb.Substring(0,2),16)
    $gg = [Convert]::ToByte($HashRgb.Substring(2,2),16)
    $bb = [Convert]::ToByte($HashRgb.Substring(4,2),16)
    [byte[]]@($rr,$gg,$bb)
}

function Get-HashFromRgbBytes {
    param([byte]$R,[byte]$G,[byte]$B)
    return ('#{0:X2}{1:X2}{2:X2}' -f $R,$G,$B)
}

function Split-ArgbHex8 {
    param([string]$ArgbHex8) # "AARRGGBB" (no '#')
    if ($ArgbHex8 -notmatch '^[0-9A-Fa-f]{8}$') {
        throw "Invalid ARGB '$ArgbHex8'. Expected 8 hex chars (AARRGGBB)."
    }
    $a = [Convert]::ToByte($ArgbHex8.Substring(0,2),16)
    $r = [Convert]::ToByte($ArgbHex8.Substring(2,2),16)
    $g = [Convert]::ToByte($ArgbHex8.Substring(4,2),16)
    $b = [Convert]::ToByte($ArgbHex8.Substring(6,2),16)
    [byte[]]@($a,$r,$g,$b)
}

function Join-ArgbHex8 {
    param([byte]$A,[byte]$R,[byte]$G,[byte]$B)
    return ('{0:X2}{1:X2}{2:X2}{3:X2}' -f $A,$R,$G,$B)
}

#endregion

#region Color math & selection

function Get-EuclideanSq {
    param(
        [byte]$R1,[byte]$G1,[byte]$B1,
        [byte]$R2,[byte]$G2,[byte]$B2
    )
    # squared distance in RGB space
    $dr = [int]$R1 - [int]$R2
    $dg = [int]$G1 - [int]$G2
    $db = [int]$B1 - [int]$B2
    return ($dr*$dr + $dg*$dg + $db*$db)
}

function Clamp-Byte {
    param([int]$x)
    if ($x -lt 0) { return [byte]0 }
    if ($x -gt 255) { return [byte]255 }
    return [byte]$x
}

# Generates up to 26*radius candidates per radius step by stepping along axis and diagonal unit directions.
# This gives a good approximation for "closest" in Euclidean space without exploring all 16M colors.
function Find-ClosestUnusedColor {
    param(
        [byte]$TR,[byte]$TG,[byte]$TB,
        [System.Collections.Generic.HashSet[string]]$UsedRgbSet # contains "rrggbb" lower-case
    )

    $targetKey = ('{0:X2}{1:X2}{2:X2}' -f $TR,$TG,$TB).ToLowerInvariant()
    if (-not $UsedRgbSet.Contains($targetKey)) {
        # Target is free — use it
        return ,$TR, $TG, $TB
    }

    # Directions = all non-zero combinations of (-1,0,1) for R,G,B (26 directions)
    $dirs = @()
    foreach ($dx in -1..1) {
        foreach ($dy in -1..1) {
            foreach ($dz in -1..1) {
                if ($dx -eq 0 -and $dy -eq 0 -and $dz -eq 0) { continue }
                $dirs += ,@($dx,$dy,$dz)
            }
        }
    }

    $best = $null
    $bestDist = [int]::MaxValue

    for ($rad = 1; $rad -le 255; $rad++) {
        foreach ($d in $dirs) {
            $rCand = Clamp-Byte([int]$TR + $rad * $d[0])
            $gCand = Clamp-Byte([int]$TG + $rad * $d[1])
            $bCand = Clamp-Byte([int]$TB + $rad * $d[2])

            $key = ('{0:X2}{1:X2}{2:X2}' -f $rCand,$gCand,$bCand).ToLowerInvariant()
            if ($UsedRgbSet.Contains($key)) { continue }

            $dist = Get-EuclideanSq -R1 $rCand -G1 $gCand -B1 $bCand -R2 $TR -G2 $TG -B2 $TB
            if ($dist -lt $bestDist) {
                $bestDist = $dist
                $best = @([byte]$rCand,[byte]$gCand,[byte]$bCand)
                if ($bestDist -eq 0) { break } # exact (shouldn't happen, already checked), but just in case
            }
        }
        if ($best) { break }
    }

    if (-not $best) {
        throw "Failed to find an unused color (this is extremely unlikely)."
    }

    return $best
}

#endregion

#region XML scanning & transformation

function Get-ThemeElements {
    param([xml]$XmlDoc)
    # Return all <Background> and <Foreground> nodes anywhere in the document
    # Using SelectNodes with XPath that matches both names.
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
    # .vstheme files typically have no namespace for these nodes; if they do,
    # this naive approach may need enhancement. For the common case, the below works:
    $bg = $XmlDoc.SelectNodes('//Background')
    $fg = $XmlDoc.SelectNodes('//Foreground')
    @($bg + $fg)
}

function Build-UsedColorSet {
    param([System.Collections.ArrayList]$Elements)
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($el in $Elements) {
        $type = $el.GetAttribute('Type')
        if ($type -ne 'CT_RAW') { continue }
        $src = $el.GetAttribute('Source')
        if ([string]::IsNullOrWhiteSpace($src)) { continue }
        try {
            $argb = Split-ArgbHex8 $src
        } catch { continue } # skip malformed
        $rgbKey = ('{0:X2}{1:X2}{2:X2}' -f $argb[1],$argb[2],$argb[3]).ToLowerInvariant()
        [void]$set.Add($rgbKey)
    }
    return $set
}

function Replace-MatchingColors {
    param(
        [System.Collections.ArrayList]$Elements,
        [byte[]]$SourceRgb,   # [R,G,B]
        [byte[]]$TargetRgb,   # [R,G,B]
        [System.Collections.Generic.HashSet[string]]$UsedRgbSet
    )

    $replacedCount = 0

    foreach ($el in $Elements) {
        $type = $el.GetAttribute('Type')
        if ($type -ne 'CT_RAW') { continue }

        $src = $el.GetAttribute('Source')
        if ([string]::IsNullOrWhiteSpace($src)) { continue }

        try {
            $argb = Split-ArgbHex8 $src
        } catch { continue }

        $a = $argb[0]; $r = $argb[1]; $g = $argb[2]; $b = $argb[3]

        # Compare RGB against SourceColor
        if ($r -eq $SourceRgb[0] -and $g -eq $SourceRgb[1] -and $b -eq $SourceRgb[2]) {
            # Pick closest unused to TargetRgb
            $bestRgb = Find-ClosestUnusedColor -TR $TargetRgb[0] -TG $TargetRgb[1] -TB $TargetRgb[2] -UsedRgbSet $UsedRgbSet
            $newR = [byte]$bestRgb[0]; $newG = [byte]$bestRgb[1]; $newB = [byte]$bestRgb[2]

            $newArgb = Join-ArgbHex8 -A $a -R $newR -G $newG -B $newB
            $el.SetAttribute('Source', $newArgb)

            # Track newly used color to prevent duplicates in subsequent replacements
            $key = ('{0:X2}{1:X2}{2:X2}' -f $newR,$newG,$newB).ToLowerInvariant()
            [void]$UsedRgbSet.Add($key)

            $replacedCount++
        }
        else {
            # No replacement; leave as is
            continue
        }
    }

    return $replacedCount
}

#endregion

#region Main

try {
    if (-not (Test-Path -LiteralPath $InputFile)) {
        throw "Input file not found: $InputFile"
    }

    $srcRgb = Get-RgbBytesFromHash $SourceColor
    $tgtRgb = Get-RgbBytesFromHash $TargetColor

    [xml]$xml = Get-Content -LiteralPath $InputFile -Raw -ErrorAction Stop

    # Collect relevant elements
    $elements = [System.Collections.ArrayList](Get-ThemeElements -XmlDoc $xml)

    # Build initial used set from CT_RAW nodes (original colors only)
    $used = Build-UsedColorSet -Elements $elements

    # Replace matching colors; as we replace, add the new RGBs to the used set to ensure uniqueness
    $count = Replace-MatchingColors -Elements $elements -SourceRgb $srcRgb -TargetRgb $tgtRgb -UsedRgbSet $used

    # Ensure output directory exists
    $outDir = Split-Path -Path $OutputFile -Parent
    if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir | Out-Null
    }

    $xml.Save($OutputFile)

    Write-Host ("Done. Replaced {0} occurrence(s) of {1} with colors based on {2}. Output: {3}" -f $count,$SourceColor,$TargetColor,$OutputFile)
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}

#endregion
