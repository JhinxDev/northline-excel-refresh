$ErrorActionPreference='Stop'
$root=(Resolve-Path "$PSScriptRoot/..").Path
$excel=$null;$book=$null;$checks=@()
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3
 if($excel.Workbooks.Count -ne 0){throw 'Expected a new empty Excel instance'}
 $book=$excel.Workbooks.Open("$root/northline-refresh.xlsx",0,$true)
 if($book.Queries.Item('DataRoot').Formula.Trim() -ne '"SET_YOUR_FIXTURES_FOLDER"'){throw 'Release DataRoot is not portable'}
 $checks+=@{check='portable_DataRoot';passed=$true}
 foreach($name in @('AllRows','Accepted','Exceptions','Monthly','Reconciliation')) {
  if($book.Queries.Item($name).Formula.Trim() -ne (Get-Content "$PSScriptRoot/$name.m" -Raw).Trim()){throw "Embedded source mismatch: $name"}
  $checks+=@{check="embedded_source/$name";passed=$true}
 }
 $expected=Get-Content "$root/evidence/expected.json" -Raw|ConvertFrom-Json
 $r=$book.Worksheets.Item('Reconciliation')
 for($i=7;$i -le 14;$i++) {
  $key=$r.Cells.Item($i,2).Value2;$value=$r.Cells.Item($i,3).Value2
  if($value -ne $expected.october.$key){throw "Saved value mismatch: $key"}
  $checks+=@{check="saved_values/$key";passed=$true;actual=$value}
 }
 @{read_only_reopen=$true;queries_executed_by_previous_receipt=$true;checks=$checks} | ConvertTo-Json -Depth 6 | Set-Content "$root/evidence/saved-workbook-readback.json"
} finally {
 if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
 if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
