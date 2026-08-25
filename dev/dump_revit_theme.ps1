# Decode Revit's own theme dictionaries, outside Revit.
#
# This is where the values in GUI/RevitTheme.py come from. UIFramework.dll is
# loaded with Assembly.LoadFrom, its UIFramework.g.resources stream is read, and
# the BAML entries are turned back into live objects with Baml2006Reader. No
# Revit session is needed and nothing is guessed.
#
#   powershell -STA -NoProfile -File dev/dump_revit_theme.ps1 -Out revit_theme.json
#
# Useful entries in the dump:
#   themes/themecolorlight.baml        90 named colours - the dialog palette
#   themes/themecolordark.baml         the same 90, dark
#   themes/applicationthemebase.baml    the CLIENT AREA + chrome colours - this is
#                                       where the real dialog background lives
#   themes/childwindowcontrolstyles.baml  the dialog style kit + its metrics
#
# themes/applicationthemefull{light,dark}.baml will NOT load here: they merge
# other dictionaries by pack URI, which needs an Application. Their values live
# in the themecolor* files above anyway.
#
# Re-run this after a Revit upgrade and diff against the palette before
# assuming anything moved.

param(
  [string]$RevitDir = "C:\Program Files\Autodesk\Revit 2026",
  [string]$Out      = "revit_theme.json"
)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xaml

$ErrorActionPreference = 'Stop'

$onResolve = [System.ResolveEventHandler]{
  param($s, $e)
  $simple = ($e.Name -split ',')[0]
  $path = Join-Path $RevitDir ($simple + '.dll')
  if (Test-Path $path) { return [Reflection.Assembly]::LoadFrom($path) }
  return $null
}
[AppDomain]::CurrentDomain.add_AssemblyResolve($onResolve)

$asm = [Reflection.Assembly]::LoadFrom((Join-Path $RevitDir 'UIFramework.dll'))
$rr  = New-Object System.Resources.ResourceReader($asm.GetManifestResourceStream('UIFramework.g.resources'))

$wanted = @(
  'themes/applicationthemefulllight.baml',
  'themes/applicationthemefulldark.baml',
  'themes/themecolorlight.baml',
  'themes/themecolordark.baml',
  'themes/applicationthemebase.baml',
  'themes/childwindowcontrolstyles.baml'
)

$result = @{}

$en = $rr.GetEnumerator()
while ($en.MoveNext()) {
  $name = [string]$en.Key
  if ($wanted -notcontains $name) { continue }
  $stream = $en.Value
  if (-not ($stream -is [System.IO.Stream])) {
    Write-Output ("{0}: value is {1}, not a stream" -f $name, $stream.GetType().Name)
    continue
  }
  try {
    $stream.Position = 0
    $reader = New-Object System.Windows.Baml2006.Baml2006Reader($stream)
    $writer = New-Object System.Xaml.XamlObjectWriter($reader.SchemaContext)
    [System.Xaml.XamlServices]::Transform($reader, $writer)
    $dict = $writer.Result
  } catch {
    Write-Output ("FAILED {0}: {1}" -f $name, $_.Exception.Message)
    continue
  }

  $bag = @{}
  foreach ($k in $dict.Keys) {
    $key = [string]$k
    $v = $null
    try { $v = $dict[$k] } catch { continue }
    if ($v -is [Windows.Media.SolidColorBrush]) {
      $bag[$key] = $v.Color.ToString()
    } elseif ($v -is [Windows.Media.Color]) {
      $bag[$key] = $v.ToString()
    } elseif ($v -ne $null) {
      $bag[$key] = ('<' + $v.GetType().Name + '>')
    }
  }
  $result[$name] = $bag
  Write-Output ("{0}: {1} entries" -f $name, $bag.Count)
}

$result | ConvertTo-Json -Depth 4 | Set-Content -Path $Out -Encoding UTF8
Write-Output ("written -> " + $Out)
