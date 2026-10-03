param([switch]$VerifyExisting)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path "$PSScriptRoot/..").Path
$excel = $null; $book = $null
$checks = [System.Collections.Generic.List[object]]::new()
function Check($name,$actual,$expected) {
 if ("$actual" -ne "$expected") { throw "$name expected [$expected], got [$actual]" }
 $checks.Add([pscustomobject]@{check=$name;actual=$actual;expected=$expected;passed=$true})
}
function Refresh-And-Check($stage,$expected) {
 foreach ($name in @('Accepted','Exceptions','Monthly','Reconciliation')) {
  $table=$book.Worksheets.Item($name).ListObjects.Item(1)
  $table.QueryTable.Refresh($false) | Out-Null
 }
 $excel.CalculateUntilAsyncQueriesDone(); $excel.CalculateFullRebuild()
 $recon=$book.Worksheets.Item('Reconciliation')
 for ($i=7;$i -le 16;$i++) {
  $key=$recon.Cells.Item($i,2).Value2; $value=$recon.Cells.Item($i,3).Value2
  $want=if($key -in @('row_difference','revenue_difference_cents')){0}else{$expected.$key}
  Check "$stage/$key" $value $want
 }
 $monthly=$book.Worksheets.Item('Monthly').ListObjects.Item(1)
 Check "$stage/month_count" $monthly.ListRows.Count @($expected.months.PSObject.Properties).Count
 for($i=1;$i -le $monthly.ListRows.Count;$i++) {
  $month=$monthly.DataBodyRange.Cells.Item($i,1).Text
  foreach($pair in @(@('orders',2),@('revenue_cents',3),@('cost_cents',4))) {
   Check "$stage/$month/$($pair[0])" $monthly.DataBodyRange.Cells.Item($i,$pair[1]).Value2 $expected.months.$month.($pair[0])
  }
 }
 $exceptions=$book.Worksheets.Item('Exceptions').ListObjects.Item(1)
 Check "$stage/visible_exceptions" $exceptions.ListRows.Count $expected.exception_rows
 if($expected.exception_rows -gt 0) {
  $actual=@(); for($i=1;$i -le $exceptions.ListRows.Count;$i++) {
   $actual+= "$($exceptions.DataBodyRange.Cells.Item($i,3).Value2)|$($exceptions.DataBodyRange.Cells.Item($i,4).Value2)"
  }
  Check "$stage/exception_reasons" (($actual|Sort-Object)-join ',') (($expected.exceptions|ForEach-Object {"$($_.order_id)|$($_.issues)"}|Sort-Object)-join ',')
 }
}
try {
 $expected=Get-Content "$root/evidence/expected.json" -Raw|ConvertFrom-Json
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.AutomationSecurity=3
 if($excel.Workbooks.Count -ne 0){throw 'Expected a new empty Excel instance'}
 $book=$excel.Workbooks.Open("$root/northline-refresh.xlsx",0,$false)
 if($VerifyExisting) {
  Check 'portable_first_run/DataRoot_placeholder' $book.Queries.Item('DataRoot').Formula.Trim() '"SET_YOUR_FIXTURES_FOLDER"'
  $book.Queries.Item('DataRoot').Formula=('"'+($root.Replace('\','/')+'/fixtures')+'"')
 }
 if(-not $VerifyExisting) {
  if(Test-Path "$root/fixtures/monthly/2026-10.csv"){throw 'October already exists. Use -VerifyExisting or an isolated fresh copy.'}
  [void]$book.Queries.Add('DataRoot',('"'+($root.Replace('\','/')+'/fixtures')+'"'))
  foreach($name in @('AllRows','Accepted','Exceptions','Monthly','Reconciliation')) {
   [void]$book.Queries.Add($name,(Get-Content "$PSScriptRoot/$name.m" -Raw))
  }
  foreach($name in @('Accepted','Exceptions','Monthly','Reconciliation')) {
   $sheet=$book.Worksheets.Item($name)
   $conn='OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location='+$name+';Extended Properties=""'
   $table=$sheet.ListObjects.Add(0,$conn,$null,1,$sheet.Range('B6'))
   $table.Name='Northline'+$name; $table.TableStyle='TableStyleLight1'
   $table.QueryTable.CommandType=2; $table.QueryTable.CommandText="SELECT * FROM [$name]"
   if($name -eq 'Exceptions') { $table.QueryTable.CommandText='SELECT [source_file], [source_row], [order_id], [issues], [date], [channel], [sku], [quantity], [unit_price_cents], [unit_cost_cents] FROM [Exceptions]' }
   $table.QueryTable.BackgroundQuery=$false; $table.QueryTable.AdjustColumnWidth=$false
  }
  Refresh-And-Check 'six_month_baseline' $expected.baseline
  Copy-Item -LiteralPath "$root/fixtures/incoming/2026-10.csv" -Destination "$root/fixtures/monthly/2026-10.csv"
 }
 Refresh-And-Check 'added_october' $expected.october
 Refresh-And-Check 'repeat_refresh' $expected.october
 $masterPath=Join-Path $root 'fixtures/products.csv'
 $originalMaster=[IO.File]::ReadAllText($masterPath)
 try {
  foreach($kind in @('empty_master_sku','whitespace_master_sku','schema_drift')) {
   $badMaster=switch($kind) {
    'empty_master_sku' { $originalMaster -replace '(?m)^BG01,',',' }
    'whitespace_master_sku' { $originalMaster -replace '(?m)^BG01,','   ,' }
    'schema_drift' { $originalMaster -replace '^sku,','unexpected_key,' }
   }
   [IO.File]::WriteAllText($masterPath,$badMaster,[Text.UTF8Encoding]::new($false))
   $rejected=$false
   try {
    $ok=$book.Worksheets.Item('Accepted').ListObjects.Item(1).QueryTable.Refresh($false)
    if(-not $ok){$rejected=$true}
   } catch { $rejected=$true }
   Check "$kind/refresh_rejected" $rejected $true
  }
 } finally { [IO.File]::WriteAllText($masterPath,$originalMaster,[Text.UTF8Encoding]::new($false)) }
 Refresh-And-Check 'restored_master' $expected.october
 $monthly=$book.Worksheets.Item('Monthly')
 $moneyFormat='#'+$excel.ThousandsSeparator+'##0'+$excel.DecimalSeparator+'00" CAD"'
 $monthly.Range('F7:G100').NumberFormatLocal=$moneyFormat
 $monthly.Range('H7:H100').NumberFormatLocal=('0'+$excel.DecimalSeparator+'0%')
 $monthly.Range('B:B').ColumnWidth=14; $monthly.Range('C:H').ColumnWidth=19
 $monthly.PageSetup.PrintArea='B2:L40'
 $exceptions=$book.Worksheets.Item('Exceptions')
 $exceptions.Range('B:B').ColumnWidth=20; $exceptions.Range('C:C').ColumnWidth=12
 $exceptions.Range('D:D').ColumnWidth=23; $exceptions.Range('E:E').ColumnWidth=35
 $exceptions.Range('E7:E2000').WrapText=$true
 $exceptions.PageSetup.PrintArea='B2:E19'
 $recon=$book.Worksheets.Item('Reconciliation')
 $recon.Range('B:B').ColumnWidth=39; $recon.Range('C:C').ColumnWidth=20
 $recon.Range('C7:C16').NumberFormatLocal=('#'+$excel.ThousandsSeparator+'##0')
 $recon.PageSetup.PrintArea='B2:F18'
 if(-not $VerifyExisting) {
  $chart=$monthly.ChartObjects().Add(15,230,790,220).Chart
  $chart.ChartType=51
  $series=$chart.SeriesCollection().NewSeries(); $series.Name='Revenue (CAD)'
  $series.XValues=$monthly.Range('B7:B13');$series.Values=$monthly.Range('F7:F13')
  $series.Format.Fill.ForeColor.RGB=6384663
  $chart.HasTitle=$true; $chart.ChartTitle.Text='Revenue from accepted orders (CAD)'
  $chart.HasLegend=$false; $chart.Axes(2).TickLabels.NumberFormat='0'
  $chart.Axes(2).MinimumScale=0
  $monthly.PageSetup.PrintArea='B2:L40'
 }
 foreach($name in @('Accepted','Exceptions','Monthly','Reconciliation')) {
  $header=$book.Worksheets.Item($name).ListObjects.Item(1).HeaderRowRange
  $header.Interior.Color=6384663; $header.Font.Color=16777215
 }
 $book.Worksheets.Item('Overview').Activate()
 $book.Save()
 foreach($name in @('Overview','Monthly','Exceptions','Reconciliation')) {
  $book.Worksheets.Item($name).ExportAsFixedFormat(0,"$root/evidence/$name.pdf")
 }
 $receiptName=if($VerifyExisting){'native-excel-reverification.json'}else{'native-excel-verification.json'}
 @{version='1.1.0';excel_version=$excel.Version;native_power_query=$true;operator_exercise='Add October CSV, synchronous query refresh, compare independent oracle, repeat refresh';checks=$checks} | ConvertTo-Json -Depth 8 | Set-Content "$root/evidence/$receiptName"
} finally {
 if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
 if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
