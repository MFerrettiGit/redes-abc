<#  build.ps1  —  Curva Redes ABC
    Lê o .xlsx exportado da query (query\curva_redes_abc.sql) e gera dados\redes.js.
    Uso:  powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -Xlsx "C:\...\curva_redes.xlsx"
    Colunas esperadas (cabeçalho, em qualquer ordem):
      rede, redeCod, anoMes, prodCod, prodDesc, marca, grupoProduto, valor, qtd
#>
param(
  [Parameter(Mandatory=$true)][string]$Xlsx,
  [string]$Root = "C:\Users\COMPRASD\redes-abc",
  [int]$MesesJanela = 12
)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

if(-not (Test-Path $Xlsx)){ throw "Arquivo não encontrado: $Xlsx" }

# ---------- leitura do xlsx (sem Excel) ----------
$zip=[System.IO.Compression.ZipFile]::OpenRead($Xlsx)
function Read-Entry($name){ $e=$zip.Entries | Where-Object { $_.FullName -eq $name }; if(-not $e){return $null}; $r=New-Object IO.StreamReader($e.Open()); $t=$r.ReadToEnd(); $r.Close(); $t }

# shared strings
$shared=New-Object System.Collections.ArrayList
$sst=Read-Entry 'xl/sharedStrings.xml'
if($sst){ [xml]$sx=$sst; foreach($si in $sx.sst.si){ [void]$shared.Add([string]$si.InnerText) } }

# descobre a 1a planilha
$wb=Read-Entry 'xl/workbook.xml'
$sheetPath='xl/worksheets/sheet1.xml'
$rowsData=New-Object System.Collections.ArrayList
function ColToIdx($ref){ $c=($ref -replace '[0-9]','' ); $n=0; foreach($ch in $c.ToCharArray()){ $n=$n*26+([int][char]$ch-64) }; return $n-1 }

$sheetXml=Read-Entry $sheetPath
$rd=[System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($sheetXml)))
$curRow=$null; $curType=$null; $curRef=$null; $inV=$false
while($rd.Read()){
  if($rd.NodeType -eq 'Element' -and $rd.Name -eq 'row'){ $curRow=@{} }
  elseif($rd.NodeType -eq 'Element' -and $rd.Name -eq 'c'){ $curRef=$rd.GetAttribute('r'); $curType=$rd.GetAttribute('t') }
  elseif($rd.NodeType -eq 'Element' -and $rd.Name -eq 'v'){
    $val=$rd.ReadElementContentAsString()
    if($curType -eq 's'){ $val=[string]$shared[[int]$val] }
    $curRow[(ColToIdx $curRef)]=$val
  }
  elseif($rd.NodeType -eq 'Element' -and $rd.Name -eq 't' -and $curType -eq 'inlineStr'){
    $val=$rd.ReadElementContentAsString(); $curRow[(ColToIdx $curRef)]=$val
  }
  elseif($rd.NodeType -eq 'EndElement' -and $rd.Name -eq 'row'){ [void]$rowsData.Add($curRow) }
}
$rd.Close(); $zip.Dispose()

if($rowsData.Count -lt 2){ throw "Planilha sem dados." }

# ---------- mapeia cabeçalho ----------
$hdr=$rowsData[0]
$col=@{}
foreach($k in $hdr.Keys){ $col[($hdr[$k].Trim().ToLower())]=$k }
function Cval($row,$name){ $i=$col[$name.ToLower()]; if($i -ne $null -and $row.ContainsKey($i)){ return $row[$i] }; return '' }
foreach($need in 'rede','redecod','anomes','prodcod','proddesc','valor'){
  if(-not $col.ContainsKey($need)){ throw "Coluna '$need' não encontrada no cabeçalho. Colunas: $($col.Keys -join ', ')" }
}

# ---------- agrega ----------
$mesesSet=New-Object System.Collections.Generic.HashSet[string]
$redes=@{}
for($i=1;$i -lt $rowsData.Count;$i++){
  $r=$rowsData[$i]
  $rede=(Cval $r 'rede').Trim(); $rcod=(Cval $r 'redecod').Trim(); $am=(Cval $r 'anomes').Trim()
  if(-not $rede -or -not $am){ continue }
  $pcod=(Cval $r 'prodcod').Trim(); $pdesc=(Cval $r 'proddesc').Trim()
  $marca=(Cval $r 'marca').Trim(); $grupo=(Cval $r 'grupoproduto').Trim()
  $valRaw=(Cval $r 'valor'); $val=0.0; [void][double]::TryParse(($valRaw -replace ',','.'),[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$val)
  [void]$mesesSet.Add($am)
  if(-not $redes.ContainsKey($rcod)){ $redes[$rcod]=@{ nome=$rede; prods=@{} } }
  $pk=$pcod
  if(-not $redes[$rcod].prods.ContainsKey($pk)){ $redes[$rcod].prods[$pk]=@{ desc=$pdesc; marca=$marca; grupo=$grupo; m=@{} } }
  $pm=$redes[$rcod].prods[$pk].m
  if(-not $pm.ContainsKey($am)){ $pm[$am]=0.0 }
  $pm[$am]+=$val
}

# Janela deterministica: 12 meses FECHADOS + o mes ATUAL (parcial) = 13 meses.
$hoje=[datetime]::Today
$primeiroMes=(Get-Date -Year $hoje.Year -Month $hoje.Month -Day 1).AddMonths(-12)
$meses=@(); for($i=0;$i -lt 13;$i++){ $meses+=$primeiroMes.AddMonths($i).ToString('yyyy-MM') }
$mesParcial=(Get-Date -Year $hoje.Year -Month $hoje.Month -Day 1).ToString('yyyy-MM')
$mesesFechados=12

# ---------- monta objeto ----------
$redesArr=@()
foreach($rc in $redes.Keys){
  $rp=$redes[$rc]
  $prodArr=@()
  foreach($pc in $rp.prods.Keys){
    $p=$rp.prods[$pc]
    $serie=@(); $tot=0.0
    foreach($m in $meses){ $v=0.0; if($p.m.ContainsKey($m)){ $v=[math]::Round($p.m[$m],2) }; $serie+=$v; $tot+=$v }
    if($tot -eq 0){ continue }
    $prodArr+=[pscustomobject]@{ cod=$pc; desc=$p.desc; marca=$p.marca; grupo=$p.grupo; serie=$serie }
  }
  if($prodArr.Count -eq 0){ continue }
  $redesArr+=[pscustomobject]@{ cod=$rc; nome=$rp.nome; produtos=$prodArr }
}

$obj=[pscustomobject]@{
  atualizadoEm=(Get-Date -Format 'dd/MM/yyyy HH:mm')
  meses=$meses
  mesParcial=$mesParcial
  mesesFechados=$mesesFechados
  redes=$redesArr
}
$json=$obj | ConvertTo-Json -Depth 8 -Compress
$js="/* GERADO por build.ps1 em $(Get-Date -Format 'dd/MM/yyyy HH:mm') — NÃO editar à mão */`r`nwindow.REDES_RAW=$json;`r`n"
$out=Join-Path $Root 'dados\redes.js'
[IO.File]::WriteAllText($out,$js,(New-Object Text.UTF8Encoding($false)))

Write-Host "OK -> $out"
Write-Host "Redes: $($redesArr.Count) | Meses: $($meses -join ', ')"
