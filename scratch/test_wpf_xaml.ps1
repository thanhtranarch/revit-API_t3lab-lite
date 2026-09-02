Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$files = @(
    "T3Lab.extension\lib\GUI\Tools\ExportManager.xaml",
    "T3Lab.extension\lib\GUI\Tools\ManaViews.xaml",
    "T3Lab.extension\lib\GUI\Tools\ManaSheets.xaml",
    "T3Lab.extension\lib\GUI\Tools\SheetGen.xaml",
    "T3Lab.extension\lib\GUI\Tools\SplitElements.xaml",
    "T3Lab.extension\lib\GUI\Tools\T3Dialog.xaml",
    "T3Lab.extension\lib\GUI\Tools\ParameterSelector.xaml",
    "T3Lab.extension\lib\GUI\Tools\SubtypeDefinerColMap.xaml",
    "T3Lab.extension\lib\GUI\Tools\CadtoFloorLayerItem.xaml",
    "T3Lab.extension\lib\GUI\Tools\FindReplace.xaml",
    "T3Lab.extension\lib\GUI\Tools\DimText.xaml",
    "T3Lab.extension\lib\GUI\Tools\TagChecker.xaml",
    "T3Lab.extension\lib\GUI\Tools\TextToElement.xaml",
    "T3Lab.extension\lib\GUI\Tools\QuickElement.xaml",
    "T3Lab.extension\lib\GUI\Tools\ManaSelect.xaml",
    "T3Lab.extension\lib\GUI\Tools\ManaAnno.xaml",
    "T3Lab.extension\lib\GUI\Tools\AutoJoin.xaml",
    "T3Lab.extension\lib\GUI\Tools\AutoWork.xaml",
    "T3Lab.extension\lib\GUI\Tools\DoorThreshold.xaml",
    "T3Lab.extension\lib\GUI\Tools\RoomToFloor.xaml",
    "T3Lab.extension\lib\GUI\Tools\CadtoFloor.xaml",
    "T3Lab.extension\lib\GUI\Tools\CadtoWall.xaml",
    "T3Lab.extension\lib\GUI\Tools\CADtoBeam.xaml",
    "T3Lab.extension\lib\GUI\Tools\TileLayout.xaml",
    "T3Lab.extension\lib\GUI\Tools\FoundationVolume.xaml",
    "T3Lab.extension\lib\GUI\Tools\ImageToDrafting.xaml",
    "T3Lab.extension\lib\GUI\Tools\PointCloud.xaml",
    "T3Lab.extension\lib\GUI\Tools\PropertyLine.xaml",
    "T3Lab.extension\lib\GUI\Tools\CADToElements.xaml",
    "T3Lab.extension\lib\GUI\Tools\AutoDimension.xaml",
    "T3Lab.extension\lib\GUI\Tools\AdvancedViewManagerBatchRename.xaml",
    "T3Lab.extension\lib\GUI\Tools\UIStandardShowcase.xaml"
)

$hasError = $false

foreach ($file in $files) {
    Write-Host "Testing $file..." -NoNewline
    try {
        # Read file, remove event handlers like Click="...", TextChanged="...", SelectionChanged="...", etc.
        # which require code-behind to instantiate via XamlReader
        $raw = [System.IO.File]::ReadAllText($file)
        $clean = [System.Text.RegularExpressions.Regex]::Replace(
            $raw,
            '\b(Click|TextChanged|SelectionChanged|Checked|Unchecked|MouseDoubleClick|PreviewMouseLeftButtonDown|MouseLeftButtonUp|SizeChanged|DropDownClosed|ContentRendered|Drop|KeyDown)="[^"]*"',
            ''
        )
        $stringReader = New-Object System.IO.StringReader($clean)
        $xmlReader = [System.Xml.XmlReader]::Create($stringReader)
        $obj = [System.Windows.Markup.XamlReader]::Load($xmlReader)
        Write-Host " [PASS] Loaded successfully in WPF!" -ForegroundColor Green
    }
    catch {
        $hasError = $true
        Write-Host " [FAIL] ERROR:" -ForegroundColor Red
        Write-Host $_.Exception.ToString() -ForegroundColor Red
    }
}

if ($hasError) {
    exit 1
} else {
    Write-Host "`nALL XAML FILES PASSED REAL WPF RUNTIME LOAD TEST!" -ForegroundColor Green
    exit 0
}
