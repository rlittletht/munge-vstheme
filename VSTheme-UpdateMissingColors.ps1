param(
    [Parameter(Mandatory=$true)]
    [string]$OriginalTheme,
    
    [Parameter(Mandatory=$true)]
    [string]$UpdatedTheme,
    
    [Parameter(Mandatory=$false)]
    [string]$NewTheme
)

# If NewTheme path not provided, put it in same directory as OriginalTheme
if ([string]::IsNullOrWhiteSpace($NewTheme)) {
    $originalFullPath = Resolve-Path $OriginalTheme
    $originalDir = Split-Path -Path $originalFullPath -Parent
    $NewTheme = Join-Path $originalDir "MergedTheme.vstheme"
    Write-Host "NewTheme not specified. Using: $NewTheme"
}

# Validate file extensions
$files = @($OriginalTheme, $UpdatedTheme, $NewTheme)
foreach ($file in $files) {
    if ($file -notmatch '\.vstheme$') {
        Write-Error "File '$file' must have a .vstheme extension"
        exit 1
    }
}

# Validate that source files exist
if (-not (Test-Path $OriginalTheme)) {
    Write-Error "OriginalTheme file not found: $OriginalTheme"
    exit 1
}

if (-not (Test-Path $UpdatedTheme)) {
    Write-Error "UpdatedTheme file not found: $UpdatedTheme"
    exit 1
}

# Load XML files
Write-Host "Loading theme files..."
$originalXml = New-Object System.Xml.XmlDocument
$originalXml.PreserveWhitespace = $true
$originalXml.Load((Resolve-Path $OriginalTheme))

$updatedXml = New-Object System.Xml.XmlDocument
$updatedXml.PreserveWhitespace = $true
$updatedXml.Load((Resolve-Path $UpdatedTheme))

# Create a hash set of color names from the original theme
$originalColors = @{}
foreach ($category in $originalXml.SelectNodes("//Category")) {
    foreach ($color in $category.SelectNodes("Color")) {
        $colorName = $color.GetAttribute("Name")
        if ($colorName) {
            $originalColors[$colorName] = $true
        }
    }
}

Write-Host "Found $($originalColors.Count) colors in original theme"

# Find new colors in the updated theme
$newColorsFound = 0
$colorsToAdd = @{}

foreach ($category in $updatedXml.SelectNodes("//Category")) {
    $categoryName = $category.GetAttribute("Name")
    
    foreach ($color in $category.SelectNodes("Color")) {
        $colorName = $color.GetAttribute("Name")
        
        # If this color doesn't exist in the original theme, track it
        if ($colorName -and -not $originalColors.ContainsKey($colorName)) {
            if (-not $colorsToAdd.ContainsKey($categoryName)) {
                $colorsToAdd[$categoryName] = @()
            }
            $colorsToAdd[$categoryName] += $color
            $newColorsFound++
        }
    }
}

Write-Host "Found $newColorsFound new colors in updated theme"

if ($newColorsFound -eq 0) {
    Write-Host "No new colors to add. Creating copy of original theme..."
    Copy-Item $OriginalTheme $NewTheme -Force
    Write-Host "New theme created: $NewTheme"
    exit 0
}

# Start with a copy of the original theme
$newXml = $originalXml.Clone()
$newXml.PreserveWhitespace = $true

# Add new colors to the appropriate categories
$colorsAdded = 0
foreach ($categoryName in $colorsToAdd.Keys) {
    # Find or create the category in the new theme
    $category = $newXml.SelectSingleNode("//Category[@Name='$categoryName']")
    
    if (-not $category) {
        Write-Host "Creating new category: $categoryName"
        # Create the category if it doesn't exist
        $themesNode = $newXml.SelectSingleNode("//Themes")
        if (-not $themesNode) {
            Write-Error "Could not find Themes node in XML structure"
            exit 1
        }
        
        $category = $newXml.CreateElement("Category")
        $category.SetAttribute("Name", $categoryName)
        $category.SetAttribute("GUID", "{" + [guid]::NewGuid().ToString() + "}")
        $themesNode.AppendChild($category) | Out-Null
    }
    
    # Add each new color to this category
    foreach ($color in $colorsToAdd[$categoryName]) {
        $importedColor = $newXml.ImportNode($color, $true)
        $category.AppendChild($importedColor) | Out-Null
        $colorsAdded++
        Write-Host "  Added color: $($color.GetAttribute('Name')) to category: $categoryName"
    }
}

# Save the new theme with proper encoding to preserve entities
Write-Host "`nSaving new theme to: $NewTheme"
$settings = New-Object System.Xml.XmlWriterSettings
$settings.Encoding = [System.Text.Encoding]::UTF8
$settings.Indent = $true
$settings.IndentChars = "  "
$settings.NewLineChars = "`r`n"
$settings.NewLineHandling = [System.Xml.NewLineHandling]::Replace

$writer = [System.Xml.XmlWriter]::Create($NewTheme, $settings)
try {
    $newXml.Save($writer)
} finally {
    $writer.Close()
}

Write-Host "Successfully added $colorsAdded colors to the new theme!"
