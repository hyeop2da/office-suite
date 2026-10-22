# Word에서 DOCX 복사본을 읽기 전용으로 열어 쪽 수, 문단, 표, 머리글·바닥글, 글상자 글자를 JSON으로 남긴다.
# 안전 규칙: 이미 실행 중인 Word가 있으면 시작하지 않는다(사용자 문서 보호). 직접 시작한 인스턴스만 종료한다.
# 사용: powershell -File tools/verify-word.ps1 <docx 폴더> <출력 json>
param([string]$Dir, [string]$Out)
if (Get-Process WINWORD -ErrorAction SilentlyContinue) { Write-Error 'Word가 이미 실행 중입니다. 사용자 문서 보호를 위해 중단합니다.'; exit 2 }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$word.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: 매크로 실행 금지
$results = @()
try {
  foreach ($f in Get-ChildItem $Dir -Filter *.docx | Sort-Object Name) {
    $r = [ordered]@{ file = $f.Name }
    try {
      # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
      $doc = $word.Documents.Open($f.FullName, $false, $true, $false)
      $r.pages = $doc.ComputeStatistics(2)
      $r.paragraphs = @($doc.Paragraphs | ForEach-Object { $_.Range.Text.TrimEnd("`r", "`a", [char]7) })
      $r.tables = $doc.Tables.Count
      $r.cells = @(); foreach ($t in $doc.Tables) { foreach ($c in $t.Range.Cells) { $r.cells += $c.Range.Text.TrimEnd("`r", "`a", [char]7) } }
      $r.header = @($doc.Sections(1).Headers(1).Range.Paragraphs | ForEach-Object { $_.Range.Text.TrimEnd("`r") })
      $r.footer = @($doc.Sections(1).Footers(1).Range.Paragraphs | ForEach-Object { $_.Range.Text.TrimEnd("`r") })
      $r.shapes = @(); foreach ($s in $doc.Shapes) { try { if ($s.TextFrame.HasText) { $r.shapes += $s.TextFrame.TextRange.Text.TrimEnd("`r") } } catch {} }
      $r.compat = $doc.CompatibilityMode
      $doc.Close($false)
    } catch { $r.error = $_.Exception.Message }
    $results += $r
  }
} finally { $word.Quit([ref]$false) }
$results | ConvertTo-Json -Depth 5 | Out-File $Out -Encoding utf8
"wrote $Out"
