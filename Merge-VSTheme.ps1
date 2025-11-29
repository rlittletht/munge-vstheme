param(
    [Parameter(Mandatory=$true)]
    [string]$ThemeFile,
    
    [Parameter(Mandatory=$true)]
    [string]$Changes,
    
    [Parameter(Mandatory=$false)]
    [string]$TargetFile
)

# Function to add .vstheme extension if not present
function Get-ThemeFilePath {
    param([string]$Path)
    
    if ([System.IO.Path]::HasExtension($Path)) {
        return $Path
    }
    return "$Path.vstheme"
}

# Function to resolve file path
function Resolve-FilePath {
    param([string]$Path)
    
    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }
    return Join-Path -Path (Get-Location) -ChildPath $Path
}

# Resolve file paths
$themeFilePath = Resolve-FilePath (Get-ThemeFilePath $ThemeFile)
$changesFilePath = Resolve-FilePath (Get-ThemeFilePath $Changes)

# Determine target file path
if ([string]::IsNullOrWhiteSpace($TargetFile)) {
    $targetFilePath = $themeFilePath
} else {
    $targetFilePath = Resolve-FilePath (Get-ThemeFilePath $TargetFile)
}

# Validate input files exist
if (-not (Test-Path $themeFilePath)) {
    Write-Error "Theme file not found: $themeFilePath"
    exit 1
}

if (-not (Test-Path $changesFilePath)) {
    Write-Error "Changes file not found: $changesFilePath"
    exit 1
}

Write-Host "Loading theme file: $themeFilePath"
Write-Host "Loading changes file: $changesFilePath"
Write-Host "Target file: $targetFilePath"

# Load XML files with PreserveWhitespace to maintain formatting
$themeXml = New-Object System.Xml.XmlDocument
$themeXml.PreserveWhitespace = $true
$themeXml.Load($themeFilePath)

$changesXml = New-Object System.Xml.XmlDocument
$changesXml.PreserveWhitespace = $true
$changesXml.Load($changesFilePath)

# Function to find or create a Category node
function Get-OrCreateCategory {
    param(
        [System.Xml.XmlDocument]$XmlDoc,
        [System.Xml.XmlElement]$ThemeNode,
        [string]$CategoryName,
        [string]$CategoryGuid
    )
    
    # Try to find existing category by GUID first, then by Name
    $category = $ThemeNode.SelectSingleNode("Category[@GUID='$CategoryGuid']")
    
    if (-not $category) {
        $category = $ThemeNode.SelectSingleNode("Category[@Name='$CategoryName']")
    }
    
    if (-not $category) {
        Write-Host "  Creating new category: $CategoryName"
        $category = $XmlDoc.CreateElement("Category")
        $category.SetAttribute("Name", $CategoryName)
        $category.SetAttribute("GUID", $CategoryGuid)
        [void]$ThemeNode.AppendChild($category)
    }
    
    return $category
}

# Function to merge or update a Color node
function Merge-ColorNode {
    param(
        [System.Xml.XmlDocument]$XmlDoc,
        [System.Xml.XmlElement]$CategoryNode,
        [System.Xml.XmlElement]$SourceColorNode
    )
    
    $colorName = $SourceColorNode.GetAttribute("Name")
    
    # Find existing color node
    $existingColor = $CategoryNode.SelectSingleNode("Color[@Name='$colorName']")
    
    if ($existingColor) {
        Write-Host "    Updating color: $colorName"
        # Remove the old node
        [void]$CategoryNode.RemoveChild($existingColor)
    } else {
        Write-Host "    Adding new color: $colorName"
    }
    
    # Import and append the new color node
    $importedNode = $XmlDoc.ImportNode($SourceColorNode, $true)
    [void]$CategoryNode.AppendChild($importedNode)
}

# Get the Theme node from the theme file
$themeNode = $themeXml.SelectSingleNode("//Theme")

if (-not $themeNode) {
    Write-Error "Could not find Theme node in theme file"
    exit 1
}

# Process each category in the changes file
$categoriesProcessed = 0
$colorsProcessed = 0

foreach ($changeCategory in $changesXml.SelectNodes("//Category")) {
    $categoryName = $changeCategory.GetAttribute("Name")
    $categoryGuid = $changeCategory.GetAttribute("GUID")
    
    Write-Host "Processing category: $categoryName"
    $categoriesProcessed++
    
    # Get or create the category in the theme
    $targetCategory = Get-OrCreateCategory -XmlDoc $themeXml -ThemeNode $themeNode -CategoryName $categoryName -CategoryGuid $categoryGuid
    
    # Merge each color from the changes into the target category
    foreach ($colorNode in $changeCategory.SelectNodes("Color")) {
        Merge-ColorNode -XmlDoc $themeXml -CategoryNode $targetCategory -SourceColorNode $colorNode
        $colorsProcessed++
    }
}

# Save the merged XML with proper settings to preserve entities
$settings = New-Object System.Xml.XmlWriterSettings
$settings.Indent = $true
$settings.IndentChars = "  "
$settings.NewLineChars = "`r`n"
$settings.Encoding = [System.Text.UTF8Encoding]::new($false) # UTF-8 without BOM
$settings.OmitXmlDeclaration = $false

Write-Host "`nSaving merged theme to: $targetFilePath"

$writer = [System.Xml.XmlWriter]::Create($targetFilePath, $settings)
try {
    $themeXml.Save($writer)
} finally {
    $writer.Close()
}

Write-Host "Successfully merged theme file!"
Write-Host "  Categories processed: $categoriesProcessed"
Write-Host "  Colors processed: $colorsProcessed"
