# PowerPoint에서 PPTX 복사본을 읽기 전용(창 없이)으로 열어 슬라이드·도형·표·노트 글자를 JSON으로 남긴다.
# 안전 규칙: 이미 실행 중인 PowerPoint가 있으면 시작하지 않는다. 직접 시작한 인스턴스만 종료한다.
param([string]$Dir, [string]$Out)
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { Write-Error 'PowerPoint가 이미 실행 중입니다. 중단합니다.'; exit 2 }
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.DisplayAlerts = 1   # ppAlertsNone
$ppt.AutomationSecurity = 3
$results = @()
try {
  foreach ($f in Get-ChildItem $Dir -Filter *.pptx | Sort-Object Name) {
    $r = [ordered]@{ file = $f.Name }
    try {
      # Open(FileName, ReadOnly, Untitled, WithWindow)
      $p = $ppt.Presentations.Open($f.FullName, -1, 0, 0)
      $r.slideCount = $p.Slides.Count
      $slides = @()
      foreach ($s in $p.Slides) {
        $sl = [ordered]@{ index = $s.SlideIndex; shapes = @(); cells = @(); notes = @() }
        foreach ($sh in $s.Shapes) {
          if ($sh.HasTable) { foreach ($row in $sh.Table.Rows) { foreach ($c in $row.Cells) { $sl.cells += $c.Shape.TextFrame.TextRange.Text } } }
          elseif ($sh.HasTextFrame -and $sh.TextFrame.HasText) { $sl.shapes += $sh.TextFrame.TextRange.Text }
        }
        foreach ($sh in $s.NotesPage.Shapes) { if ($sh.HasTextFrame -and $sh.TextFrame.HasText) { $sl.notes += $sh.TextFrame.TextRange.Text } }
        $slides += $sl
      }
      $r.slides = $slides
      $p.Close()
    } catch { $r.error = $_.Exception.Message }
    $results += $r
  }
} finally { $ppt.Quit() }
$results | ConvertTo-Json -Depth 6 | Out-File $Out -Encoding utf8
"wrote $Out"
