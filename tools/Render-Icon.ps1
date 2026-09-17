param(
    [string]$SvgPath = (Join-Path $PSScriptRoot '..\src\MeetingAlarm.App\Assets\MeetingAlarm.svg'),
    [string]$PngPath = (Join-Path $PSScriptRoot '..\src\MeetingAlarm.App\Assets\MeetingAlarm.png')
)
$ErrorActionPreference = 'Stop'
# Run with Windows PowerShell -STA; WPF supplies native vector rasterization.
Add-Type -AssemblyName PresentationCore, WindowsBase
[xml]$svg = Get-Content $SvgPath -Raw
$culture = [System.Globalization.CultureInfo]::InvariantCulture
function Number($value, $fallback = 0) {
    if ([string]::IsNullOrEmpty($value)) { return $fallback }
    return [double]::Parse($value, $culture)
}
$gradients = @{}
foreach ($gradient in $svg.svg.defs.linearGradient) {
    $brush = [System.Windows.Media.LinearGradientBrush]::new()
    $brush.MappingMode = [System.Windows.Media.BrushMappingMode]::Absolute
    $brush.StartPoint = [System.Windows.Point]::new((Number $gradient.x1), (Number $gradient.y1))
    $brush.EndPoint = [System.Windows.Point]::new((Number $gradient.x2), (Number $gradient.y2))
    foreach ($stop in $gradient.stop) {
        $color = [System.Windows.Media.ColorConverter]::ConvertFromString($stop.'stop-color')
        $brush.GradientStops.Add([System.Windows.Media.GradientStop]::new($color, (Number $stop.offset)))
    }
    $gradients[$gradient.id] = $brush
}
function Brush($paint, $opacity = 1) {
    if (!$paint -or $paint -eq 'none') { return $null }
    if ($paint -match '^url\(#(.+)\)$') { $brush = $gradients[$Matches[1]].Clone() }
    else { $brush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($paint)) }
    $brush.Opacity = $opacity
    return $brush
}
$visual = [System.Windows.Media.DrawingVisual]::new()
$drawing = $visual.RenderOpen()
try {
    foreach ($element in $svg.svg.ChildNodes) {
        if ($element.LocalName -notin @('rect','circle','path')) { continue }
        $fill = Brush $element.GetAttribute('fill') (Number $element.GetAttribute('fill-opacity') 1)
        $stroke = Brush $element.GetAttribute('stroke') (Number $element.GetAttribute('stroke-opacity') 1)
        $pen = $null
        if ($stroke) {
            $pen = [System.Windows.Media.Pen]::new($stroke, (Number $element.GetAttribute('stroke-width') 1))
            if ($element.GetAttribute('stroke-linecap') -eq 'round') {
                $pen.StartLineCap = $pen.EndLineCap = [System.Windows.Media.PenLineCap]::Round
            }
            if ($element.GetAttribute('stroke-linejoin') -eq 'round') { $pen.LineJoin = [System.Windows.Media.PenLineJoin]::Round }
        }
        switch ($element.LocalName) {
            'rect' {
                $rectangle = [System.Windows.Rect]::new((Number $element.x), (Number $element.y), (Number $element.width), (Number $element.height))
                $radius = Number $element.rx
                $drawing.DrawRoundedRectangle($fill, $pen, $rectangle, $radius, $radius)
            }
            'circle' {
                $point = [System.Windows.Point]::new((Number $element.cx), (Number $element.cy))
                $drawing.DrawEllipse($fill, $pen, $point, (Number $element.r), (Number $element.r))
            }
            'path' { $drawing.DrawGeometry($fill, $pen, [System.Windows.Media.Geometry]::Parse($element.d)) }
        }
    }
}
finally { $drawing.Close() }
$bitmap = [System.Windows.Media.Imaging.RenderTargetBitmap]::new(512, 512, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
$bitmap.Render($visual)
$encoder = [System.Windows.Media.Imaging.PngBitmapEncoder]::new()
$encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
$stream = [System.IO.File]::Create($PngPath)
try { $encoder.Save($stream) }
finally { $stream.Dispose() }
Write-Host "Rendered $PngPath from the original SVG"
