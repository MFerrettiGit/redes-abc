<#  rodar.ps1  —  Curva Redes ABC
    Conecta no SQL Server (credencial do Cofre do Windows), roda a query
    query\curva_redes_abc.sql e gera dados\redes.js direto (sem xlsx).
    Uso: powershell -ExecutionPolicy Bypass -File scripts\rodar.ps1
    Mesmo acesso da automacao NaoCompra (target Ferretti-LancamentosSQL).
#>
param(
  [string]$Root = "C:\Users\COMPRASD\redes-abc",
  [int]$MesesJanela = 12
)
$ErrorActionPreference='Stop'
$SERVER="189.126.153.75,2270"
$DATABASE="CO136Y_160463_PR_PD"
$CRED_TARGET="Ferretti-LancamentosSQL"
$QUERY=Join-Path $Root "query\curva_redes_abc.sql"

# ---- credencial do Cofre ----
if (-not ([System.Management.Automation.PSTypeName]'CredMan').Type) {
  Add-Type -Namespace '' -Name CredMan -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("advapi32.dll", CharSet=System.Runtime.InteropServices.CharSet.Unicode, SetLastError=true)]
public static extern bool CredRead(string target, int type, int flags, out IntPtr credential);
[System.Runtime.InteropServices.DllImport("advapi32.dll")] public static extern void CredFree(IntPtr cred);
[System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
public struct CREDENTIAL { public int Flags; public int Type; public IntPtr TargetName; public IntPtr Comment;
  public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten; public int CredentialBlobSize; public IntPtr CredentialBlob;
  public int Persist; public int AttributeCount; public IntPtr Attributes; public IntPtr TargetAlias; public IntPtr UserName; }
public static string GetUser(string target){ IntPtr p; if(!CredRead(target,1,0,out p)) return null; var c=(CREDENTIAL)System.Runtime.InteropServices.Marshal.PtrToStructure(p,typeof(CREDENTIAL)); var u=System.Runtime.InteropServices.Marshal.PtrToStringUni(c.UserName); CredFree(p); return u; }
public static string GetPass(string target){ IntPtr p; if(!CredRead(target,1,0,out p)) return null; var c=(CREDENTIAL)System.Runtime.InteropServices.Marshal.PtrToStructure(p,typeof(CREDENTIAL)); var s=c.CredentialBlobSize>0?System.Runtime.InteropServices.Marshal.PtrToStringUni(c.CredentialBlob,c.CredentialBlobSize/2):null; CredFree(p); return s; }
'@
}
$sqlUser=[CredMan]::GetUser($CRED_TARGET)
$sqlPass=[CredMan]::GetPass($CRED_TARGET)
if(-not $sqlPass){ throw "Credencial '$CRED_TARGET' nao encontrada no Cofre." }
Write-Host "Credencial OK (user: $sqlUser)"

# ---- conecta e roda ----
$cs="Server=$SERVER;Database=$DATABASE;User Id=$sqlUser;Password=$sqlPass;Encrypt=True;TrustServerCertificate=True;Connect Timeout=60"
$sql=Get-Content $QUERY -Raw -Encoding UTF8
$conn=New-Object System.Data.SqlClient.SqlConnection($cs)
$conn.Open(); Write-Host "Conectado ao SQL Server $SERVER"
$cmd=$conn.CreateCommand(); $cmd.CommandText=$sql; $cmd.CommandTimeout=900
$rd=$cmd.ExecuteReader()
$inv=[System.Globalization.CultureInfo]::InvariantCulture

# indices das colunas por nome
$ord=@{}; for($i=0;$i -lt $rd.FieldCount;$i++){ $ord[$rd.GetName($i).ToLower()]=$i }
foreach($need in 'rede','redecod','anomes','prodcod','proddesc','valor'){ if(-not $ord.ContainsKey($need)){ throw "Coluna '$need' nao veio no resultado." } }
function Cell($col){ $i=$ord[$col]; if($rd.IsDBNull($i)){return ''}; return ("$($rd.GetValue($i))").Trim() }

$mesesSet=New-Object System.Collections.Generic.HashSet[string]
$redes=@{}
$n=0
while($rd.Read()){
  $rede=Cell 'rede'; $rcod=Cell 'redecod'; $am=Cell 'anomes'
  if(-not $rede -or -not $am){ continue }
  $pcod=Cell 'prodcod'; $pdesc=Cell 'proddesc'; $marca=Cell 'marca'; $grupo=Cell 'grupoproduto'
  $valRaw=Cell 'valor'; $val=0.0; [void][double]::TryParse(($valRaw -replace ',','.'),[Globalization.NumberStyles]::Any,$inv,[ref]$val)
  [void]$mesesSet.Add($am)
  if(-not $redes.ContainsKey($rcod)){ $redes[$rcod]=@{ nome=$rede; prods=@{} } }
  if(-not $redes[$rcod].prods.ContainsKey($pcod)){ $redes[$rcod].prods[$pcod]=@{ desc=$pdesc; marca=$marca; grupo=$grupo; m=@{} } }
  $pm=$redes[$rcod].prods[$pcod].m
  if(-not $pm.ContainsKey($am)){ $pm[$am]=0.0 }
  $pm[$am]+=$val
  $n++
  if($n % 50000 -eq 0){ Write-Host "... $n linhas" }
}
$rd.Close(); $conn.Close()
Write-Host "Linhas lidas: $n | Redes: $($redes.Count)"
if($n -lt 1){ throw "Nenhuma linha retornada." }

# janela: 12 meses mais recentes presentes
# Janela deterministica: 12 meses FECHADOS + o mes ATUAL (parcial) = 13 meses.
$hoje=[datetime]::Today
$primeiroMes=(Get-Date -Year $hoje.Year -Month $hoje.Month -Day 1).AddMonths(-12)
$meses=@(); for($i=0;$i -lt 13;$i++){ $meses+=$primeiroMes.AddMonths($i).ToString('yyyy-MM') }
$mesParcial=(Get-Date -Year $hoje.Year -Month $hoje.Month -Day 1).ToString('yyyy-MM')
$mesesFechados=12

$redesArr=@()
foreach($rc in $redes.Keys){
  $rp=$redes[$rc]; $prodArr=@()
  foreach($pc in $rp.prods.Keys){
    $p=$rp.prods[$pc]; $serie=@(); $tot=0.0
    foreach($m in $meses){ $v=0.0; if($p.m.ContainsKey($m)){ $v=[math]::Round($p.m[$m],2) }; $serie+=$v; $tot+=$v }
    if($tot -le 0){ continue }
    $prodArr+=[pscustomobject]@{ cod=$pc; desc=$p.desc; marca=$p.marca; grupo=$p.grupo; serie=$serie }
  }
  if($prodArr.Count -eq 0){ continue }
  $redesArr+=[pscustomobject]@{ cod=$rc; nome=$rp.nome; produtos=$prodArr }
}

$obj=[pscustomobject]@{ atualizadoEm=(Get-Date -Format 'dd/MM/yyyy HH:mm'); meses=$meses; mesParcial=$mesParcial; mesesFechados=$mesesFechados; redes=$redesArr }
$json=$obj | ConvertTo-Json -Depth 8 -Compress
$js="/* GERADO por rodar.ps1 em $(Get-Date -Format 'dd/MM/yyyy HH:mm') - dados reais do SQL Server */`r`nwindow.REDES_RAW=$json;`r`n"
$out=Join-Path $Root 'dados\redes.js'
[IO.File]::WriteAllText($out,$js,(New-Object Text.UTF8Encoding($false)))
Write-Host "OK -> $out"
Write-Host "Redes publicaveis: $($redesArr.Count) | Meses: $($meses -join ', ')"
