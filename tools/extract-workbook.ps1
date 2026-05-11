param(
  [string]$WorkbookPath = "C:\Users\Laptop\Desktop\PFE SUSU\Bilan et synthéses des arrêts du JANVIER 2026 VF.xlsx",
  [string]$OutputPath = ".\app-data.js"
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

function ConvertTo-ColumnNumber([string]$Column) {
  $n = 0
  foreach ($ch in $Column.ToUpperInvariant().ToCharArray()) {
    $n = ($n * 26) + ([int][char]$ch - [int][char]'A' + 1)
  }
  return $n
}

function ConvertFrom-ExcelSerial($Value) {
  if ($null -eq $Value -or "$Value" -eq "") { return $null }
  try {
    return ([datetime]::FromOADate([double]$Value)).ToString("yyyy-MM-ddTHH:mm:ss")
  } catch {
    return $null
  }
}

function ConvertFrom-ExcelDurationHours($Value) {
  if ($null -eq $Value -or "$Value" -eq "") { return $null }
  try {
    return [math]::Round(([double]$Value) * 24, 4)
  } catch {
    return $null
  }
}

function Get-Numeric($Value) {
  if ($null -eq $Value -or "$Value" -eq "") { return $null }
  try { return [double]$Value } catch { return $null }
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)

function Read-ZipEntry([string]$Name) {
  $entry = $script:zip.GetEntry($Name)
  if ($null -eq $entry) { return $null }
  $reader = [IO.StreamReader]::new($entry.Open())
  $text = $reader.ReadToEnd()
  $reader.Close()
  return $text
}

$sharedXml = Read-ZipEntry "xl/sharedStrings.xml"
$sharedStrings = @()
if ($sharedXml) {
  foreach ($m in [regex]::Matches($sharedXml, "<si.*?</si>", "Singleline")) {
    $parts = [regex]::Matches($m.Value, "<t[^>]*>(.*?)</t>", "Singleline") | ForEach-Object {
      [System.Net.WebUtility]::HtmlDecode($_.Groups[1].Value)
    }
    $sharedStrings += (($parts -join "") -replace "\s+", " ").Trim()
  }
}

function Get-CellRawValue($Cell, $NsManager) {
  $vNode = $Cell.SelectSingleNode("x:v", $NsManager)
  if ($null -eq $vNode) { return "" }
  $value = $vNode.InnerText
  $type = $Cell.GetAttribute("t")
  if ($type -eq "s" -and $value -ne "") {
    $idx = [int]$value
    if ($idx -ge 0 -and $idx -lt $script:sharedStrings.Count) {
      return $script:sharedStrings[$idx]
    }
  }
  return $value
}

function Get-CellFormula($Cell, $NsManager) {
  $fNode = $Cell.SelectSingleNode("x:f", $NsManager)
  if ($null -eq $fNode) { return $null }
  return (($fNode.InnerText -replace "\s+", " ") -replace "_xlfn\.", "")
}

function Get-RowMap($Row, $NsManager) {
  $map = @{}
  foreach ($cell in $Row.SelectNodes("x:c", $NsManager)) {
    $addr = $cell.GetAttribute("r")
    if ($addr -match "^([A-Z]+)(\d+)$") {
      $map[$Matches[1]] = [pscustomobject]@{
        Address = $addr
        Value   = Get-CellRawValue $cell $NsManager
        Formula = Get-CellFormula $cell $NsManager
      }
    }
  }
  return $map
}

function Get-Value($Map, [string]$Column) {
  if ($Map.ContainsKey($Column)) { return $Map[$Column].Value }
  return ""
}

function Get-FormulaValue($Map, [string]$Column) {
  if ($Map.ContainsKey($Column)) { return $Map[$Column].Formula }
  return $null
}

[xml]$workbookXml = Read-ZipEntry "xl/workbook.xml"
[xml]$relsXml = Read-ZipEntry "xl/_rels/workbook.xml.rels"
$relationshipTargets = @{}
foreach ($rel in $relsXml.Relationships.Relationship) {
  $relationshipTargets[$rel.Id] = $rel.Target
}

$nsUri = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
$sheets = @{}
$sheetInfos = @()
$formulas = @()

foreach ($sheet in $workbookXml.workbook.sheets.sheet) {
  $rid = $sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
  $path = "xl/" + $relationshipTargets[$rid]
  [xml]$sheetXml = Read-ZipEntry $path
  $nsm = [Xml.XmlNamespaceManager]::new($sheetXml.NameTable)
  $nsm.AddNamespace("x", $nsUri)

  $cells = $sheetXml.SelectNodes("//x:c", $nsm)
  $formulaCells = $sheetXml.SelectNodes("//x:c[x:f]", $nsm)
  $maxRow = 0
  $maxCol = 0

  foreach ($cell in $cells) {
    if ($cell.GetAttribute("r") -match "^([A-Z]+)(\d+)$") {
      $rowNo = [int]$Matches[2]
      $colNo = ConvertTo-ColumnNumber $Matches[1]
      if ($rowNo -gt $maxRow) { $maxRow = $rowNo }
      if ($colNo -gt $maxCol) { $maxCol = $colNo }
    }
  }

  $sheetInfos += [pscustomobject]@{
    name         = $sheet.name
    dimension    = $sheetXml.worksheet.dimension.ref
    cells        = $cells.Count
    formulaCount = $formulaCells.Count
    maxRow       = $maxRow
    maxCol       = $maxCol
  }

  foreach ($cell in $formulaCells) {
    $formulas += [pscustomobject]@{
      sheet   = $sheet.name
      address = $cell.GetAttribute("r")
      formula = Get-CellFormula $cell $nsm
      cached  = Get-CellRawValue $cell $nsm
    }
  }

  $sheets[$sheet.name] = [pscustomobject]@{
    Xml = $sheetXml
    Ns  = $nsm
  }
}

$bilan = $sheets["Bilan"]
$events = @()
foreach ($row in $bilan.Xml.SelectNodes("//x:sheetData/x:row", $bilan.Ns)) {
  $rowNo = [int]$row.GetAttribute("r")
  if ($rowNo -lt 46) { continue }
  $map = Get-RowMap $row $bilan.Ns
  $startRaw = Get-Value $map "D"
  $endRaw = Get-Value $map "E"
  $nature = (Get-Value $map "C").Trim()
  $description = (Get-Value $map "G").Trim()
  if ($startRaw -eq "" -and $endRaw -eq "" -and $nature -eq "" -and $description -eq "") { continue }
  $durationRaw = Get-Value $map "F"
  $durationHours = ConvertFrom-ExcelDurationHours $durationRaw
  if ($null -eq $durationHours -and $startRaw -ne "" -and $endRaw -ne "") {
    $durationHours = [math]::Round((([double]$endRaw - [double]$startRaw) * 24), 4)
  }
  $events += [pscustomobject]@{
    id            = "AR-$rowNo"
    row           = $rowNo
    sectionKey    = (Get-Value $map "A").Trim()
    subEquipment  = (Get-Value $map "B").Trim()
    family        = $nature
    start         = ConvertFrom-ExcelSerial $startRaw
    end           = ConvertFrom-ExcelSerial $endRaw
    durationHours = $durationHours
    description   = $description
    assignment    = (Get-Value $map "H").Trim()
    quality       = (Get-Value $map "I").Trim()
    destination   = (Get-Value $map "J").Trim()
  }
}

$familiesSheet = $sheets["Familles arrêts"]
$families = @()
foreach ($row in $familiesSheet.Xml.SelectNodes("//x:sheetData/x:row", $familiesSheet.Ns)) {
  $rowNo = [int]$row.GetAttribute("r")
  if ($rowNo -lt 5) { continue }
  $map = Get-RowMap $row $familiesSheet.Ns
  $name = (Get-Value $map "C").Trim()
  if ($name -eq "") { continue }
  $families += [pscustomobject]@{
    name     = $name
    examples = (Get-Value $map "D").Trim()
  }
}

$tonnageSheet = $sheets["Tonnage"]
$qualities = @("K02", "K09s", "K10", "K12", "K20", "K62", "K03", "K08", "K01")
$tonnage = @()
foreach ($row in $tonnageSheet.Xml.SelectNodes("//x:sheetData/x:row", $tonnageSheet.Ns)) {
  $rowNo = [int]$row.GetAttribute("r")
  if ($rowNo -lt 7 -or $rowNo -gt 37) { continue }
  $map = Get-RowMap $row $tonnageSheet.Ns
  $dateRaw = Get-Value $map "B"
  $pesage = @{}
  $draft = @{}
  $pesageCols = @("C","D","E","F","G","H","I","J","K")
  $draftCols = @("P","Q","R","S","T","U","V","W","X")
  for ($i = 0; $i -lt $qualities.Count; $i++) {
    $pesage[$qualities[$i]] = Get-Numeric (Get-Value $map $pesageCols[$i])
    $draft[$qualities[$i]] = Get-Numeric (Get-Value $map $draftCols[$i])
  }
  $hasData = ($pesage.Values | Where-Object { $null -ne $_ }).Count -gt 0 -or ($draft.Values | Where-Object { $null -ne $_ }).Count -gt 0
  if (-not $hasData -and $dateRaw -eq "") { continue }
  $tonnage += [pscustomobject]@{
    day         = ConvertFrom-ExcelSerial $dateRaw
    daySerial   = Get-Numeric $dateRaw
    pesage      = $pesage
    pesageTotal = Get-Numeric (Get-Value $map "M")
    draft       = $draft
    draftTotal  = Get-Numeric (Get-Value $map "Z")
  }
}

$trainsSheet = $sheets["Trains"]
$trains = @()
foreach ($row in $trainsSheet.Xml.SelectNodes("//x:sheetData/x:row", $trainsSheet.Ns)) {
  $rowNo = [int]$row.GetAttribute("r")
  if ($rowNo -lt 5 -or $rowNo -gt 35) { continue }
  $map = Get-RowMap $row $trainsSheet.Ns
  $count = Get-Numeric (Get-Value $map "B")
  if ($null -eq $count) { continue }
  $trains += [pscustomobject]@{
    day                    = ConvertFrom-ExcelSerial (Get-Value $map "A")
    trains                 = $count
    wagons                 = Get-Numeric (Get-Value $map "C")
    durationHours          = ConvertFrom-ExcelDurationHours (Get-Value $map "D")
    averageHours           = ConvertFrom-ExcelDurationHours (Get-Value $map "E")
    silos                  = Get-Numeric (Get-Value $map "F")
    tonnageDA              = Get-Numeric (Get-Value $map "G")
    tonnageDB              = Get-Numeric (Get-Value $map "H")
    tonnageBascule         = Get-Numeric (Get-Value $map "I")
    totalTonnage           = Get-Numeric (Get-Value $map "J")
    affectationHours       = ConvertFrom-ExcelDurationHours (Get-Value $map "K")
    delayHours             = ConvertFrom-ExcelDurationHours (Get-Value $map "L")
    cadenceTph             = Get-Numeric (Get-Value $map "M")
    trsMaintenanceExploit  = Get-Numeric (Get-Value $map "N")
    semiWetTrains          = Get-Numeric (Get-Value $map "O")
  }
}

$shipsSheet = $sheets["Navire"]
$ships = @()
foreach ($row in $shipsSheet.Xml.SelectNodes("//x:sheetData/x:row", $shipsSheet.Ns)) {
  $rowNo = [int]$row.GetAttribute("r")
  if ($rowNo -lt 5) { continue }
  $map = Get-RowMap $row $shipsSheet.Ns
  $name = (Get-Value $map "C").Trim()
  if ($name -eq "") { continue }
  $ships += [pscustomobject]@{
    number        = Get-Numeric (Get-Value $map "A")
    berth         = (Get-Value $map "B").Trim()
    name          = $name
    quality       = (Get-Value $map "D").Trim()
    ecNumber      = (Get-Value $map "E").Trim()
    start         = ConvertFrom-ExcelSerial (Get-Value $map "F")
    end           = ConvertFrom-ExcelSerial (Get-Value $map "G")
    durationHours = ConvertFrom-ExcelDurationHours (Get-Value $map "H")
    scaleA        = Get-Numeric (Get-Value $map "I")
    scaleB        = Get-Numeric (Get-Value $map "J")
    scaleC        = Get-Numeric (Get-Value $map "K")
    scaleD        = Get-Numeric (Get-Value $map "L")
    bascule       = Get-Numeric (Get-Value $map "M")
    connaissement = Get-Numeric (Get-Value $map "N")
    gapRatio      = Get-Numeric (Get-Value $map "O")
    observation   = (Get-Value $map "P").Trim()
  }
}

$referenceSheets = @("EXPORTER","Feuil1  (3)","Feuil1  (4)","Feuil 5","Feuil 6 ","Feuil1","Feuil 7","Feuil 7 (2)")
$requests = @()
foreach ($sheetName in $referenceSheets) {
  if (-not $sheets.ContainsKey($sheetName)) { continue }
  $sheet = $sheets[$sheetName]
  foreach ($row in $sheet.Xml.SelectNodes("//x:sheetData/x:row", $sheet.Ns)) {
    $rowNo = [int]$row.GetAttribute("r")
    if ($rowNo -lt 5) { continue }
    $map = Get-RowMap $row $sheet.Ns
    $family = (Get-Value $map "C").Trim()
    $description = (Get-Value $map "G").Trim()
    $sectionKey = (Get-Value $map "A").Trim()
    $subEquipment = (Get-Value $map "B").Trim()
    if ($family -eq "" -and $description -eq "" -and $sectionKey -eq "" -and $subEquipment -eq "") { continue }
    $requests += [pscustomobject]@{
      source       = $sheetName.Trim()
      row          = $rowNo
      sectionKey   = $sectionKey
      subEquipment = $subEquipment
      family       = $family
      start        = ConvertFrom-ExcelSerial (Get-Value $map "D")
      end          = ConvertFrom-ExcelSerial (Get-Value $map "E")
      durationHours = ConvertFrom-ExcelDurationHours (Get-Value $map "F")
      description  = $description
      assignment   = (Get-Value $map "H").Trim()
      quality      = (Get-Value $map "I").Trim()
      destination  = (Get-Value $map "J").Trim()
    }
  }
}

$externalLinks = @()
foreach ($entry in $zip.Entries | Where-Object { $_.FullName -like "xl/externalLinks/_rels/*.rels" }) {
  $reader = [IO.StreamReader]::new($entry.Open())
  $text = $reader.ReadToEnd()
  $reader.Close()
  foreach ($m in [regex]::Matches($text, 'Target="([^"]+)"')) {
    $externalLinks += [System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value)
  }
}

$payload = [pscustomobject]@{
  generatedAt   = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
  sourceWorkbook = $WorkbookPath
  reportContext = [pscustomobject]@{
    site = "Direction des Embarquements - Port de Casablanca"
    process = "Poste de Commande, manutention phosphate, chargement navires et dechargement trains"
    method = "DMAIC / Measure-Analyze-Improve"
    needs = @(
      "Saisie et suivi des anomalies et arrets des installations",
      "Elaboration automatique des bilans journaliers, mensuels et annuels",
      "Suivi des KPI de manutention: tonnage, cadence, TRS, TRG, disponibilite",
      "Traçabilite SMQE des evenements, qualites, navires, trains et affectations",
      "Aide a la decision par Pareto, familles d'arrets et priorisation maintenance"
    )
  }
  sheets        = $sheetInfos
  formulas      = $formulas
  formulasCount = $formulas.Count
  families      = $families
  events        = $events
  tonnage       = $tonnage
  trains        = $trains
  ships         = $ships
  requests      = $requests
  externalLinks = $externalLinks
}

$json = $payload | ConvertTo-Json -Depth 8 -Compress
$content = "window.PFE_DATA = $json;`n"
$resolvedOutput = if ([IO.Path]::IsPathRooted($OutputPath)) { $OutputPath } else { Join-Path (Get-Location) $OutputPath }
[IO.File]::WriteAllText($resolvedOutput, $content, [System.Text.UTF8Encoding]::new($false))
$zip.Dispose()

Write-Host "Generated $resolvedOutput"
Write-Host "Events: $($events.Count), formulas: $($formulas.Count), requests: $($requests.Count), ships: $($ships.Count), trains: $($trains.Count)"
