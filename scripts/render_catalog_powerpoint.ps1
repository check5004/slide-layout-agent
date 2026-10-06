param(
    [Parameter(Mandatory=$true)][string]$Pptx,
    [Parameter(Mandatory=$true)][string]$OutputDir
)
# Optional development QA. Normal Python generation never calls this script.
$ErrorActionPreference='Stop'
$deckPath=(Resolve-Path -LiteralPath $Pptx).Path
$renderDir=[System.IO.Path]::GetFullPath($OutputDir)
if(Test-Path -LiteralPath $renderDir){
    if(@(Get-ChildItem -LiteralPath $renderDir -Force).Count -gt 0){throw 'OutputDir must be empty.'}
} else {New-Item -ItemType Directory -Path $renderDir | Out-Null}
$app=$null; $presentation=$null
$measurements=[System.Collections.Generic.List[object]]::new()
try {
    $app=New-Object -ComObject PowerPoint.Application
    $presentation=$app.Presentations.Open($deckPath,-1,0,0)
    foreach($slide in $presentation.Slides){
        $slide.Export((Join-Path $renderDir ('slide-{0:D2}.png' -f $slide.SlideIndex)),'PNG',1600,900)
        foreach($shape in $slide.Shapes){
            if($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1){
                $range=$shape.TextFrame2.TextRange
                $measurements.Add([pscustomobject]@{slide=$slide.SlideIndex;shape=$shape.Name;kind='text';
                    boxWidth=[math]::Round($shape.Width,2);boxHeight=[math]::Round($shape.Height,2);
                    textWidth=[math]::Round($range.BoundWidth,2);textHeight=[math]::Round($range.BoundHeight,2);
                    overflow=($range.BoundHeight -gt ($shape.Height+1) -or $range.BoundWidth -gt ($shape.Width+1))})
            }
            if($shape.HasTable -eq -1){
                for($r=1;$r -le $shape.Table.Rows.Count;$r++){
                    for($c=1;$c -le $shape.Table.Columns.Count;$c++){
                        $cell=$shape.Table.Cell($r,$c).Shape
                        if($cell.TextFrame.HasText -eq -1){
                            $range=$cell.TextFrame2.TextRange
                            $measurements.Add([pscustomobject]@{slide=$slide.SlideIndex;shape=$shape.Name;kind='table';row=$r;column=$c;
                                boxWidth=[math]::Round($cell.Width,2);boxHeight=[math]::Round($cell.Height,2);
                                textWidth=[math]::Round($range.BoundWidth,2);textHeight=[math]::Round($range.BoundHeight,2);
                                overflow=($range.BoundHeight -gt ($cell.Height+1) -or $range.BoundWidth -gt ($cell.Width+1))})
                        }
                    }
                }
            }
        }
    }
    $report=[pscustomobject]@{engine='Microsoft PowerPoint';version=$app.Version;slideCount=$presentation.Slides.Count;
        pptxSha256=(Get-FileHash -LiteralPath $deckPath -Algorithm SHA256).Hash.ToLower();
        rasterExport=$true;editableTextBoundaryMeasurements=$measurements;
        overflowCount=@($measurements | Where-Object {$_.overflow}).Count;
        chartReview='native chart pixels and data-label placement require separate visual review';
        visualInspection='exported; not yet visually reviewed'}
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $renderDir 'render-report.json') -Encoding utf8
    $report | Select-Object engine,version,slideCount,overflowCount | ConvertTo-Json
} finally {
    if($presentation){$presentation.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)}
    if($app){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
}
