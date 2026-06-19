$ErrorActionPreference = "Stop"
$docId = "193hm5A4YK9FczhVGXxtgo"
$apiKey = "863a5652184fa2a988f217019a3ebf751f7d3fc7"
$base = "https://onlinedata.kushkurriculum.org/api/docs/$docId"
$headers = @{ Authorization = "Bearer $apiKey"; "Content-Type" = "application/json" }

function Get-Tables {
  (Invoke-RestMethod -Method GET -Headers $headers -Uri "$base/tables").tables
}

function Get-ColumnIds($tableId) {
  (Invoke-RestMethod -Method GET -Headers $headers -Uri "$base/tables/$([uri]::EscapeDataString($tableId))/columns").columns | ForEach-Object { $_.id }
}

function Get-RecordCount($tableId) {
  $resp = Invoke-RestMethod -Method GET -Headers $headers -Uri "$base/tables/$([uri]::EscapeDataString($tableId))/records"
  return @($resp.records).Count
}

Write-Host "--- BEFORE ---"
$tables = Get-Tables
$tableIds = @($tables | ForEach-Object { $_.id })
$tableIds | Sort-Object | ForEach-Object { Write-Host "Table: $_" }

$moorCandidates = $tableIds | Where-Object { $n = ($_ -replace "[^A-Za-z0-9]","").ToLowerInvariant(); ($n -like "moordocument*") -or ($n -like "moortrustee*") } | Sort-Object

Write-Host "`nMoor candidate tables and counts:"
foreach ($t in $moorCandidates) {
  try { $count = Get-RecordCount $t; Write-Host " - $t (records: $count)" } catch { Write-Host " - $t (records: unavailable)" }
}

$keep = @("Moor_Document","Moor_Trustee")
$toDelete = $moorCandidates | Where-Object { $keep -notcontains $_ }

Write-Host "`nWill keep: $($keep -join ", ")"
Write-Host "Will delete duplicates: $($toDelete -join ", ")"

foreach ($t in $toDelete) {
  try { 
    Invoke-RestMethod -Method DELETE -Headers $headers -Uri "$base/tables/$([uri]::EscapeDataString($t))" | Out-Null
    Write-Host "Deleted table: $t" 
  } catch { 
    Write-Host "Failed deleting table $t" 
  }
}

$beneficiary = "Beneficiary"
$dupCols = @("trustEffectiveDay2","trustEffectiveMonth2","trustEffectiveYear2","employerIdentificationNumber2")
try {
  $existingCols = @(Get-ColumnIds $beneficiary)
  foreach ($c in $dupCols) {
    if ($existingCols -contains $c) {
      try { 
        Invoke-RestMethod -Method DELETE -Headers $headers -Uri "$base/tables/$([uri]::EscapeDataString($beneficiary))/columns/$([uri]::EscapeDataString($c))" | Out-Null
        Write-Host "Deleted Beneficiary column: $c" 
      } catch { 
        Write-Host "Failed deleting Beneficiary column $c" 
      }
    } else { Write-Host "Column not present (already clean): $c" }
  }
} catch { Write-Host "Failed checking Beneficiary columns" }

Write-Host "`n--- AFTER ---"
$tablesAfter = Get-Tables
$tableIdsAfter = @($tablesAfter | ForEach-Object { $_.id })
$tableIdsAfter | Sort-Object | ForEach-Object { Write-Host "Table: $_" }

Write-Host "`nPost-clean Moor tables:"
$tableIdsAfter | Where-Object { $n = ($_ -replace "[^A-Za-z0-9]","").ToLowerInvariant(); ($n -like "moordocument*") -or ($n -like "moortrustee*") } | Sort-Object | ForEach-Object { Write-Host " - $_" }

Write-Host "`nBeneficiary columns now:"
Get-ColumnIds "Beneficiary" | ForEach-Object { Write-Host " - $_" }
