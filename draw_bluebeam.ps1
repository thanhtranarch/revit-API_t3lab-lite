Add-Type -AssemblyName System.Drawing

$width = 64
$height = 64
$bitmap = New-Object System.Drawing.Bitmap($width, $height)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)

$graphics.Clear([System.Drawing.Color]::Transparent)

$blueColor = [System.Drawing.Color]::FromArgb(255, 0, 119, 215)
$brush = New-Object System.Drawing.SolidBrush($blueColor)

# Left edge
$graphics.FillRectangle($brush, 2, 2, 10, 60)
# Top edge
$graphics.FillRectangle($brush, 22, 2, 40, 10)
# Right edge
$graphics.FillRectangle($brush, 52, 2, 10, 60)
# Bottom edge
$graphics.FillRectangle($brush, 2, 52, 60, 10)
# Inner square
$graphics.FillRectangle($brush, 22, 22, 20, 20)

$savePath = "T3Lab.extension\T3Lab.tab\Support.panel\CloudLinks.stack\Bluebeam Status.urlbutton\icon.png"
$bitmap.Save($savePath, [System.Drawing.Imaging.ImageFormat]::Png)

$graphics.Dispose()
$bitmap.Dispose()
$brush.Dispose()
Write-Host "Icon saved to $savePath"
