# ================================
# qpw profile integration
# - Uses ONLY WindowManager2 (latest)
# - Win32 fallback if WM2 is unavailable
# - Resolves qpw.ps1 next to this profile
# ================================

# --- Resolve qpw.ps1 path (same folder as profile file) ---
$script:QpwPath = Join-Path (Split-Path -Parent $PROFILE) 'qpw.ps1'

if (-not (Test-Path $script:QpwPath)) {
    Write-Warning "qpw.ps1 not found next to profile: $script:QpwPath"
}

# ---------------------------------------------------------
# Win32 fallback: window titles
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# Win32 fallback: monitors (for d1/d2 completion)
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# Unified window access (WindowManager2 ONLY)
# ---------------------------------------------------------
function Get-QpwWindowTitles {
    if ("WindowManager2" -as [type]) {
        [WindowManager2]::GetWindows() |
            ForEach-Object { $_.Title } |
            Where-Object { $_ }
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

# ---------------------------------------------------------
# qpw wrapper function
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# Argument completion (qpw only)
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# OPTIONAL: TAB cycles only for qpw
# ---------------------------------------------------------
try {
    Import-Module PSReadLine -ErrorAction SilentlyContinue | Out-Null
    if (Get-Module PSReadLine) {
        Set-PSReadLineKeyHandler -Key Tab -ScriptBlock {
            $line = $null; $cursor = $null
            [Microsoft.PowerShell.PSConsoleReadLine]::GetBufferState([ref]$line,[ref]$cursor)

            if ($line -match '^\s*qpw(\s|$)') {
                [Microsoft.PowerShell.PSConsoleReadLine]::Complete()
            } else {
                [Microsoft.PowerShell.PSConsoleReadLine]::MenuComplete()
            }
        }
    }
} catch { }
