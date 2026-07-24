<#
publish.ps1 - Cria (se preciso) o repo MFerrettiGit/redes-abc, faz commit/push e ativa o GitHub Pages.
Le o token do GitHub do Gerenciador de Credenciais do Windows (git:https://github.com).
Uso: powershell -ExecutionPolicy Bypass -File scripts\publish.ps1 -Message "msg"
#>
param(
  [string]$Message = "Atualiza site Curva Redes ABC",
  [string]$RepoDir = "C:\Users\COMPRASD\redes-abc",
  [string]$Owner = "MFerrettiGit",
  [string]$Repo  = "redes-abc",
  [switch]$Private
)
$ErrorActionPreference = "Continue"
$sig = @"
using System;
using System.Runtime.InteropServices;
public class CredVaultRA {
  [DllImport("advapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
  static extern bool CredRead(string target, uint type, uint flags, out IntPtr cred);
  [StructLayout(LayoutKind.Sequential)]
  struct CREDENTIAL { public uint Flags; public uint Type; public IntPtr TargetName; public IntPtr Comment;
    public long LastWritten; public uint CredentialBlobSize; public IntPtr CredentialBlob;
    public uint Persist; public uint AttributeCount; public IntPtr Attributes; public IntPtr TargetAlias; public IntPtr UserName; }
  public static string Get(string target){ IntPtr p; if(!CredRead(target,1,0,out p)) return "FAIL";
    var c=(CREDENTIAL)Marshal.PtrToStructure(p,typeof(CREDENTIAL)); return Marshal.PtrToStringUni(c.CredentialBlob,(int)c.CredentialBlobSize/2); }
}
"@
Add-Type -TypeDefinition $sig -Language CSharp
$tok = [CredVaultRA]::Get("git:https://github.com")
if($tok -eq "FAIL"){ throw "Credencial do GitHub nao encontrada (git:https://github.com)." }
$hdr = @{ Authorization = "token $tok"; "User-Agent" = "ferretti-redes-abc"; Accept = "application/vnd.github+json" }

# 1) Garante que o repo existe
try {
  Invoke-RestMethod -Uri "https://api.github.com/repos/$Owner/$Repo" -Headers $hdr -Method Get | Out-Null
  Write-Output "Repo ja existe."
} catch {
  Write-Output "Criando repo $Owner/$Repo (private=$($Private.IsPresent)) ..."
  $body = @{ name=$Repo; private=$Private.IsPresent; description="Curva Redes ABC - ranking, evolucao e saude das redes (M. Ferretti)"; has_issues=$false; has_wiki=$false } | ConvertTo-Json
  Invoke-RestMethod -Uri "https://api.github.com/user/repos" -Headers $hdr -Method Post -Body $body -ContentType "application/json" | Out-Null
  Start-Sleep -Seconds 2
}

# 2) git init + commit + push
Set-Location $RepoDir
if(-not (Test-Path "$RepoDir\.git")){ git init | Out-Null; git branch -M main }
git add -A
$pendente = git status --porcelain
if($pendente){ git commit -m $Message | Out-Null } else { Write-Output "Nada para commitar." }

git push "https://$($Owner):$tok@github.com/$Owner/$Repo.git" main 2>$null
$pushExit = $LASTEXITCODE
Write-Output ("PUSH_EXIT=" + $pushExit)
if ($pushExit -ne 0) { throw "git push falhou (exit $pushExit)" }

# 3) Ativa GitHub Pages (branch main, raiz)
try {
  Invoke-RestMethod -Uri "https://api.github.com/repos/$Owner/$Repo/pages" -Headers $hdr -Method Get | Out-Null
  Write-Output "Pages ja configurado."
} catch {
  try {
    $pbody = @{ source = @{ branch="main"; path="/" } } | ConvertTo-Json
    Invoke-RestMethod -Uri "https://api.github.com/repos/$Owner/$Repo/pages" -Headers $hdr -Method Post -Body $pbody -ContentType "application/json" | Out-Null
    Write-Output "Pages ativado."
  } catch { Write-Output ("Aviso ao ativar Pages: " + $_.Exception.Message) }
}
Write-Output ("Site: https://" + $Owner.ToLower() + ".github.io/" + $Repo + "/")
