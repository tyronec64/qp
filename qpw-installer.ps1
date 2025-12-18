<#
Install-Qpw.ps1

What it does (interactive):
1) Create profile folders:
   - $env:USERPROFILE\Documents\PowerShell          (PowerShell 7+)
   - $env:USERPROFILE\Documents\WindowsPowerShell   (Windows PowerShell 5.1)
2) Copy qpw.ps1 into both folders (expects qpw.ps1 next to this installer script)
3) Create qpw-Microsoft.PowerShell_profile.ps1 snippet in both folders
4) Add a dot-source block into both PS5+PS7 profile files (idempotent)
5) Optionally dot-source the snippet now for the current session

Notes:
- Execution Policy can still block dot-sourcing. The installer detects and prints what to do.
- This installs WindowManager2-only (fallback to Win32 if WM2 not loaded).
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-YesNo([string]$Prompt, [bool]$DefaultYes = $true) {
    $suffix = if ($DefaultYes) { "[Y/n]" } else { "[y/N]" }
    while ($true) {
        $ans = Read-Host "$Prompt $suffix"
        if ([string]::IsNullOrWhiteSpace($ans)) { return $DefaultYes }
        switch ($ans.Trim().ToLowerInvariant()) {
            'y' { return $true }
            'yes' { return $true }
            'n' { return $false }
            'no' { return $false }
        }
    }
}

function Ensure-Dir([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Ensure-File([string]$Path) {
    $dir = Split-Path -Parent $Path
    Ensure-Dir $dir
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType File -Path $Path | Out-Null
    }
}

function Get-ExecPolicySummary {
    try {
        $pol = Get-ExecutionPolicy -List
        return $pol
    } catch {
        return $null
    }
}

function Is-ExecutionBlocked {
    # If CurrentUser or LocalMachine is Restricted/AllSigned, dot-sourcing may be blocked for unsigned scripts.
    # Actual behavior depends on scope precedence and origin marking; this is a conservative check.
    $pol = Get-ExecPolicySummary
    if ($null -eq $pol) { return $false }

    $effective = Get-ExecutionPolicy
    if ($effective -in @('Restricted')) { return $true }
    return $false
}

# -------------------------
# Targets (PS7 + PS5)
# -------------------------
$docs = Join-Path $env:USERPROFILE 'Documents'
$ps7Dir = Join-Path $docs 'PowerShell'
$ps5Dir = Join-Path $docs 'WindowsPowerShell'

$ps7Profile = Join-Path $ps7Dir 'Microsoft.PowerShell_profile.ps1'
$ps5Profile = Join-Path $ps5Dir 'Microsoft.PowerShell_profile.ps1'

$ps7Snippet = Join-Path $ps7Dir 'qpw-Microsoft.PowerShell_profile.ps1'
$ps5Snippet = Join-Path $ps5Dir 'qpw-Microsoft.PowerShell_profile.ps1'

$sourceQpw = Join-Path $PSScriptRoot 'qpw.ps1'

# -------------------------
# 4 questions (as requested)
# -------------------------
$doFolders = Read-YesNo "Create profile folders in Documents (PowerShell and WindowsPowerShell)?" $true
$doCopyQpw = Read-YesNo "Copy qpw.ps1 into both profile folders (expects qpw.ps1 next to this installer)?" $true
$doProfileEdit = Read-YesNo "Add dot-source block to BOTH PS7 and PS5 profile files (idempotent)?" $true
$doDotSourceNow = Read-YesNo "Dot-source qpw snippet now for THIS session (enables completion immediately)?" $true

# -------------------------
# 1) Create folders
# -------------------------
if ($doFolders) {
    Ensure-Dir $ps7Dir
    Ensure-Dir $ps5Dir
}

# -------------------------
# 2) Copy qpw.ps1
# -------------------------
if ($doCopyQpw) {
    if (-not (Test-Path -LiteralPath $sourceQpw)) {
        throw "qpw.ps1 not found next to installer: $sourceQpw"
    }
    Ensure-Dir $ps7Dir
    Ensure-Dir $ps5Dir
    Copy-Item -LiteralPath $sourceQpw -Destination (Join-Path $ps7Dir 'qpw.ps1') -Force
    Copy-Item -LiteralPath $sourceQpw -Destination (Join-Path $ps5Dir 'qpw.ps1') -Force
}

# -------------------------
# Snippet content (WindowManager2 only + Win32 fallback)
# -------------------------
$snippetContent = @'
# ================================
# qpw profile snippet (WindowManager2-only + Win32 fallback)
# ================================

# Resolve qpw.ps1 next to THIS snippet file
$script:QpwPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'qpw.ps1'

# --- Win32 fallback: window titles (used if WindowManager2 isn't loaded) ---
if (-not ('WinTopLevelEnum' -as [type])) {
    Add-Type @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public static class WinTopLevelEnum {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll", SetLastError=true)] static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
    static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    public static List<string> GetTitles() {
        var list = new List<string>();
        EnumWindows((h, p) => {
            if (!IsWindowVisible(h)) return true;
            int len = GetWindowTextLength(h);
            if (len == 0) return true;
            var sb = new StringBuilder(len + 1);
            GetWindowText(h, sb, sb.Capacity);
            var t = sb.ToString();
            if (!string.IsNullOrWhiteSpace(t)) list.Add(t);
            return true;
        }, IntPtr.Zero);
        return list;
    }
}
"@
}

# --- Win32 fallback: monitors (for d1/d2 completion) ---
if (-not ('WinMonEnum' -as [type])) {
    Add-Type @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public static class WinMonEnum {
    public delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData);

    [DllImport("user32.dll")]
    static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip, MonitorEnumProc lpfnEnum, IntPtr dwData);

    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int left, top, right, bottom; }

    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto)]
    public struct MONITORINFOEX {
        public int cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public uint dwFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst=32)]
        public string szDevice;
    }

    public static List<string> GetMonitorNames() {
        var names = new List<string>();
        EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero,
            (IntPtr h, IntPtr d, ref RECT r, IntPtr p) => {
                MONITORINFOEX mi = new MONITORINFOEX();
                mi.cbSize = Marshal.SizeOf(typeof(MONITORINFOEX));
                if (GetMonitorInfo(h, ref mi)) names.Add(mi.szDevice);
                return true;
            },
            IntPtr.Zero);
        return names;
    }
}
"@
}

function Get-QpwWindowTitles {
    if ("WindowManager2" -as [type]) {
        [WindowManager2]::GetWindows() | ForEach-Object { $_.Title } | Where-Object { $_ }
    } else {
        [WinTopLevelEnum]::GetTitles()
    }
}

function Get-QpwWindowCount {
    if ("WindowManager2" -as [type]) {
        ([WindowManager2]::GetWindows()).Count
    } else {
        ([WinTopLevelEnum]::GetTitles()).Count
    }
}

function qpw {
    [CmdletBinding()]
    param(
        [Parameter(Position=0)]
        [string]$SearchOrQuick,
        [Parameter(Position=1, ValueFromRemainingArguments=$true)]
        [string[]]$Rest
    )

    if (-not (Test-Path $script:QpwPath)) {
        throw "qpw.ps1 not found at: $script:QpwPath"
    }

    if ($PSBoundParameters.ContainsKey('SearchOrQuick')) {
        & $script:QpwPath $SearchOrQuick @Rest
    } else {
        & $script:QpwPath @Rest
    }
}

Register-ArgumentCompleter -CommandName qpw -ParameterName SearchOrQuick -ScriptBlock {
    param($CommandName, $ParameterName, $WordToComplete)

    try {
        $titles = Get-QpwWindowTitles

        $needleRaw = $WordToComplete
        if ($null -eq $needleRaw) { $needleRaw = "" }
        $needle = $needleRaw.Trim('"', "'")

        $rx = $null
        if (-not [string]::IsNullOrWhiteSpace($needle)) {
            try {
                $rx = New-Object System.Text.RegularExpressions.Regex(
                    $needle,
                    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
                )
            } catch {
                $rx = New-Object System.Text.RegularExpressions.Regex(
                    [System.Text.RegularExpressions.Regex]::Escape($needle),
                    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
                )
            }
        }

        foreach ($t in ($titles | Where-Object { $_ -and (-not $rx -or $rx.IsMatch($_)) } | Sort-Object -Unique)) {
            $out = if ($t -match '\s') { '"' + $t + '"' } else { $t }
            [System.Management.Automation.CompletionResult]::new($out, $out, 'ParameterValue', $t)
        }
    } catch { }
}

Register-ArgumentCompleter -CommandName qpw -ParameterName Rest -ScriptBlock {
    param($CommandName, $ParameterName, $WordToComplete)

    try {
        $wCount = Get-QpwWindowCount
        $mCount = ([WinMonEnum]::GetMonitorNames()).Count

        $wins = @()
        if ($wCount -gt 0) { $wins = 1..$wCount | ForEach-Object { "w$_" } }

        $mons = @()
        if ($mCount -gt 0) { $mons = 1..$mCount | ForEach-Object { "d$_" } }

        $dirs  = @("l","r","u","d","ul","ur","dl","dr","br","bl","tl","tr")
        $quads = @("q1","q2","q3","q4")
        $acts  = @("snap","winindex","maxindex","undo","redo")

        $all = @($wins + $mons + $dirs + $quads + $acts) | Select-Object -Unique
        $token = $WordToComplete
        if ($null -eq $token) { $token = "" }

        foreach ($s in ($all | Where-Object { $_ -like "*$token*" } | Sort-Object)) {
            [System.Management.Automation.CompletionResult]::new($s, $s, 'ParameterValue', $s)
        }
    } catch { }
}
'@

# -------------------------
# 3) Write snippets
# -------------------------
Ensure-File $ps7Snippet
Ensure-File $ps5Snippet
Set-Content -LiteralPath $ps7Snippet -Value $snippetContent -Encoding UTF8
Set-Content -LiteralPath $ps5Snippet -Value $snippetContent -Encoding UTF8

# -------------------------
# 4) Edit profiles (dot-source block, idempotent)
# -------------------------
$markerStart = '# >>> qpw install (managed) >>>'
$markerEnd   = '# <<< qpw install (managed) <<<'

function Ensure-ProfileBlock([string]$ProfilePath, [string]$SnippetPath) {
    Ensure-File $ProfilePath
    $text = Get-Content -LiteralPath $ProfilePath -Raw

    if ($text -match [regex]::Escape($markerStart)) {
        # Already installed: keep it idempotent
        return
    }

    $block = @"
$markerStart
try {
    . '$SnippetPath'
} catch {
    Write-Warning "qpw profile snippet failed to load: $SnippetPath"
}
$markerEnd
"@

    # Append with spacing
    if (-not [string]::IsNullOrWhiteSpace($text) -and -not $text.EndsWith("`r`n")) {
        Add-Content -LiteralPath $ProfilePath -Value "`r`n"
    }
    Add-Content -LiteralPath $ProfilePath -Value $block
}

if ($doProfileEdit) {
    Ensure-ProfileBlock $ps7Profile $ps7Snippet
    Ensure-ProfileBlock $ps5Profile $ps5Snippet
}

# -------------------------
# Execution policy check + optional dot-source now
# -------------------------
$blocked = Is-ExecutionBlocked
if ($blocked) {
    Write-Warning "ExecutionPolicy appears to block running .ps1 scripts in this session."
    Write-Warning "Fix (recommended): Set-ExecutionPolicy -Scope CurrentUser RemoteSigned"
    Write-Warning "Or one-off: start PowerShell with: -ExecutionPolicy Bypass"
}

if ($doDotSourceNow) {
    # Dot-source the snippet for the CURRENT session
    $snippetToLoad = if ($PSVersionTable.PSEdition -eq 'Core') { $ps7Snippet } else { $ps5Snippet }
    try {
        . $snippetToLoad
        Write-Host "Loaded qpw snippet into current session: $snippetToLoad"
    } catch {
        Write-Warning "Could not dot-source snippet (likely ExecutionPolicy): $snippetToLoad"
    }
}

Write-Host "Done."
Write-Host "PS7 profile: $ps7Profile"
Write-Host "PS5 profile: $ps5Profile"
Write-Host "PS7 qpw.ps1:  $(Join-Path $ps7Dir 'qpw.ps1')"
Write-Host "PS5 qpw.ps1:  $(Join-Path $ps5Dir 'qpw.ps1')"
