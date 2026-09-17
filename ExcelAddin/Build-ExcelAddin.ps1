param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'SpecPdfGenerator.xlam'),
    [switch]$UpdateRibbonOnly
)

$ErrorActionPreference = 'Stop'
$modulePath = Join-Path $PSScriptRoot 'SpecPdfGeneratorAddin.bas'
$ribbonPath = Join-Path $PSScriptRoot 'customUI.xml'
$assetsPath = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$importPdfImagePath = (Get-ChildItem -LiteralPath $assetsPath -File | Where-Object { $_.Name -match 'pdf-32\.png$' } | Select-Object -First 1).FullName
$configureGeneratorImagePath = (Get-ChildItem -LiteralPath $assetsPath -File | Where-Object { $_.Name -match '-32\.png$' -and $_.Name -notmatch 'pdf' } | Select-Object -First 1).FullName
$temporaryPath = Join-Path (Split-Path -Parent $OutputPath) (([IO.Path]::GetFileNameWithoutExtension($OutputPath)) + '.building.xlam')
$temporaryModulePath = Join-Path ([IO.Path]::GetTempPath()) ('SpecPdfGeneratorAddin_' + [guid]::NewGuid().ToString('N') + '_cp1251.bas')

if (-not (Test-Path -LiteralPath $modulePath) -or -not (Test-Path -LiteralPath $ribbonPath)) {
    throw 'Не найдены исходники VBA или Ribbon XML.'
}

New-Item -ItemType Directory -Path (Split-Path -Parent $OutputPath) -Force | Out-Null
if (-not (Test-Path -LiteralPath $importPdfImagePath) -or -not (Test-Path -LiteralPath $configureGeneratorImagePath)) {
    throw 'Button image files are missing.'
}
if (-not $UpdateRibbonOnly) {
    Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        # The VBE imports .bas files using the current ANSI code page.
        # Keep the source as UTF-8, but import a temporary Windows-1251 copy.
        $moduleText = [IO.File]::ReadAllText($modulePath, (New-Object Text.UTF8Encoding($false)))
        [IO.File]::WriteAllText($temporaryModulePath, $moduleText, [Text.Encoding]::GetEncoding(1251))
        [void]$workbook.VBProject.VBComponents.Import($temporaryModulePath)
        # xlOpenXMLAddIn = 55
        $workbook.SaveAs($temporaryPath, 55)
    }
    catch {
        throw "Unable to build XLAM. Enable Excel Trust access to the VBA project object model. Original error: $($_.Exception.Message)"
    }
    finally {
        if ($workbook) { $workbook.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }
        if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
        Remove-Item -LiteralPath $temporaryModulePath -Force -ErrorAction SilentlyContinue
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    # Keep the last known-good add-in until Excel has produced a complete replacement.
    Move-Item -LiteralPath $temporaryPath -Destination $OutputPath -Force
}
elseif (-not (Test-Path -LiteralPath $OutputPath)) {
    throw "Cannot update Ribbon because the add-in does not exist: $OutputPath"
}
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$archive = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    function Get-ZipText([string]$entryName) {
        $entry = $archive.GetEntry($entryName)
        if (-not $entry) { throw "ZIP entry not found: $entryName" }
        $reader = New-Object System.IO.StreamReader($entry.Open())
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    }

    function Set-ZipText([string]$entryName, [string]$text) {
        $existing = $archive.GetEntry($entryName)
        if ($existing) { $existing.Delete() }
        $entry = $archive.CreateEntry($entryName)
        $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
        try { $writer.Write($text) } finally { $writer.Dispose() }
    }

    function Set-ZipBytes([string]$entryName, [byte[]]$bytes) {
        $existing = $archive.GetEntry($entryName)
        if ($existing) { $existing.Delete() }
        $entry = $archive.CreateEntry($entryName)
        $stream = $entry.Open()
        try { $stream.Write($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
    }

    $contentTypes = [xml](Get-ZipText '[Content_Types].xml')
    $types = $contentTypes.DocumentElement
    $namespace = $contentTypes.DocumentElement.NamespaceURI
    $ribbonOverrides = @($contentTypes.SelectNodes("/*[local-name()='Types']/*[local-name()='Override' and (@PartName='/customUI/customUI.xml' or @PartName='/customUI/customUI14.xml') ]"))
    foreach ($existingOverride in $ribbonOverrides) { [void]$existingOverride.ParentNode.RemoveChild($existingOverride) }
    $override = $contentTypes.CreateElement('Override', $namespace)
    [void]$override.SetAttribute('PartName', '/customUI/customUI.xml')
    [void]$override.SetAttribute('ContentType', 'application/xml')
    [void]$types.AppendChild($override)
    $pngType = $contentTypes.SelectSingleNode("/*[local-name()='Types']/*[local-name()='Default' and @Extension='png']")
    if (-not $pngType) {
        $pngType = $contentTypes.CreateElement('Default', $namespace)
        [void]$pngType.SetAttribute('Extension', 'png')
        [void]$pngType.SetAttribute('ContentType', 'image/png')
        [void]$types.AppendChild($pngType)
    }
    Set-ZipText '[Content_Types].xml' $contentTypes.OuterXml

    $rels = [xml](Get-ZipText '_rels/.rels')
    $relationships = $rels.DocumentElement
    $relationshipNamespace = $relationships.NamespaceURI
    $ribbonRelations = @($rels.SelectNodes("/*[local-name()='Relationships']/*[local-name()='Relationship' and contains(@Type, '/ui/extensibility') ]"))
    foreach ($existingRelation in $ribbonRelations) { [void]$existingRelation.ParentNode.RemoveChild($existingRelation) }
    $ribbonRelation = $rels.CreateElement('Relationship', $relationshipNamespace)
    [void]$ribbonRelation.SetAttribute('Id', 'rIdSpecPdfGeneratorRibbon')
    [void]$ribbonRelation.SetAttribute('Type', 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility')
    [void]$ribbonRelation.SetAttribute('Target', 'customUI/customUI.xml')
    [void]$relationships.AppendChild($ribbonRelation)
    Set-ZipText '_rels/.rels' $rels.OuterXml

    $customUiRelsPath = 'customUI/_rels/customUI.xml.rels'
    if ($archive.GetEntry($customUiRelsPath)) {
        $customUiRels = [xml](Get-ZipText $customUiRelsPath)
    }
    else {
        $customUiRels = [xml]'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />'
    }
    $customUiRelationships = $customUiRels.DocumentElement
    $customUiRelationshipNamespace = $customUiRelationships.NamespaceURI
    foreach ($existingImageRelation in @($customUiRels.SelectNodes("/*[local-name()='Relationships']/*[local-name()='Relationship' and (@Id='importPdfImage' or @Id='configureGeneratorImage') ]"))) {
        [void]$existingImageRelation.ParentNode.RemoveChild($existingImageRelation)
    }
    foreach ($imageDefinition in @(@('importPdfImage', 'images/import-pdf-32.png'), @('configureGeneratorImage', 'images/configure-generator-32.png'))) {
        $imageRelation = $customUiRels.CreateElement('Relationship', $customUiRelationshipNamespace)
        [void]$imageRelation.SetAttribute('Id', $imageDefinition[0])
        [void]$imageRelation.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
        [void]$imageRelation.SetAttribute('Target', $imageDefinition[1])
        [void]$customUiRelationships.AppendChild($imageRelation)
    }
    Set-ZipText $customUiRelsPath $customUiRels.OuterXml
    Set-ZipBytes 'customUI/images/import-pdf-32.png' ([IO.File]::ReadAllBytes($importPdfImagePath))
    Set-ZipBytes 'customUI/images/configure-generator-32.png' ([IO.File]::ReadAllBytes($configureGeneratorImagePath))
    $legacyRibbon = $archive.GetEntry('customUI/customUI14.xml')
    if ($legacyRibbon) { $legacyRibbon.Delete() }
    Set-ZipText 'customUI/customUI.xml' ([IO.File]::ReadAllText($ribbonPath, [Text.Encoding]::UTF8))
}
finally {
    $archive.Dispose()
}

Write-Host "Создана надстройка: $OutputPath"
