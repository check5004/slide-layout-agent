param(
    [Parameter(Mandatory=$true)][string]$Pptx,
    [Parameter(Mandatory=$true)][string]$OutputDir
)
$ErrorActionPreference = 'Stop'
$deckPath = (Resolve-Path -LiteralPath $Pptx).Path
$renderDirectory = [System.IO.Path]::GetFullPath($OutputDir)
if(Test-Path -LiteralPath $renderDirectory){
    if(@(Get-ChildItem -LiteralPath $renderDirectory -Force).Count -gt 0){
        throw 'OutputDir must be empty; choose a new folder to preserve prior renders.'
    }
}else{New-Item -ItemType Directory -Path $renderDirectory | Out-Null}
$powerPoint = $null
$presentation = $null
$measurements = [System.Collections.Generic.List[object]]::new()
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    # ReadOnly=true, Untitled=false, WithWindow=false. No macros/external links in generated decks.
    $presentation = $powerPoint.Presentations.Open($deckPath, -1, 0, 0)
    foreach($slide in $presentation.Slides){
        $filename = 'slide-{0:D2}.png' -f $slide.SlideIndex
        $slide.Export((Join-Path $renderDirectory $filename), 'PNG', 1600, 900)
        foreach($shape in $slide.Shapes){
            if($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1){
                $range = $shape.TextFrame2.TextRange
                $measurements.Add([pscustomobject]@{
                    slide=$slide.SlideIndex; shape=$shape.Name;
                    boxWidth=[math]::Round($shape.Width,2); boxHeight=[math]::Round($shape.Height,2);
                    textWidth=[math]::Round($range.BoundWidth,2); textHeight=[math]::Round($range.BoundHeight,2);
                    overflow=($range.BoundHeight -gt ($shape.Height + 1) -or $range.BoundWidth -gt ($shape.Width + 1))
                })
            }
        }
    }
    $result = [pscustomobject]@{
        engine='Microsoft PowerPoint'; slideCount=$presentation.Slides.Count;
        rasterExport=$true; editableTextBoundaryMeasurements=$measurements;
        overflowCount=@($measurements | Where-Object {$_.overflow}).Count;
        visualInspection='required separately; exporting PNG does not mean a human reviewed it'
    }
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $renderDirectory 'render-report.json') -Encoding utf8
    $result | Select-Object engine,slideCount,rasterExport,overflowCount | ConvertTo-Json
} finally {
    if($presentation){$presentation.Close(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)}
    # Do not quit the application: COM may attach to an existing user instance.
    if($powerPoint){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
