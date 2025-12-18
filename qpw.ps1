<#
.SYNOPSIS
  QuickPath  PowerShell window placement utility.

.DESCRIPTION
  Move/resize the current or selected window across monitors using quick tokens.
  Supports Windows display numbering, grids, quadrants, optional activation,
  configurable split percentage and spacer margin.

.PARAMETER Monitors
  List monitors using Windows numbering and exit.

.PARAMETER Reload
  Force recompile of helper class to avoid stale types.

.PARAMETER Activate
  Bring the target window to foreground for this run (default: no activation).

.PARAMETER SplitPercent
  Default split percentage for l/r/t/b splits (default: 50).

.PARAMETER Spacer
  Pixel margin to apply on all sides of the final region (default: 0).

.EXAMPLE
  .\\qpw.ps1 b
  Bottom half (uses SplitPercent=50 by default).

.EXAMPLE
  .\\qpw.ps1 -SplitPercent 25 lrr
  Use 25% splits for chained left/right.

.EXAMPLE
  .\\qpw.ps1 -Spacer 8 d1l
  Move to Display 1, left region with 8px margin.

.EXAMPLE
  .\\qpw.ps1 d2l30a
  Display 2, left 30% and activate (token 'a').
#>
# qpw.ps1
# jan@mccs.nl
# Quick window placement with:
#   - Quick commands: w{N}[d{M}|sd{M}] [dirs] [grid] [quadrant] [action]
#       dirs = l/r/t/b/f/d (left/right/top/bottom/full/down; each halves region; f = full work area; d = alias of b)
#       grid = NxM:rRcC (e.g., 3x3:r2c1 or 4x4:r3c4), overrides dirs if present
#       quadrants = q1..q4 (q1=tl, q2=tr, q3=bl, q4=br)
#       actions = m|min|minimize, M|max|maximize
#   - Display token is optional; if missing, window’s current monitor is used.
#   - Token order doesn't matter: w5tl, w5d2tl, d2w5f, w5sd2rrb, w1m, w1M all work.
#   - Undo/Redo: "undo", "redo"
# CLI usage:
#   - .\qpw.ps1 w1d2dll            # DIRECT quick command (no search)
#   - .\qpw.ps1 d1f                # DIRECT quick (current foreground window → display 1 → full)
#   - .\qpw.ps1 ddl                # DIRECT quick (current window → down, down, left)
#   - .\qpw.ps1 <search> <quick>   # Search-first, then quick on single match (index optional)
#   - .\qpw.ps1 <search>           # Just filter, then interact
# upcoming changes:
# have the default split on 50% but then configure the divers in 25% or 75% so you can place a window on 75% of the screen and the command box on the 25%
# divers from 100% is defau;lt 50% but then when specified if you say -divers 75 means when you do t of top it will place it on 75% on the screen
#
# add the force close and close options
# add the monitoring order correctly *order but the logical resolution increases (look at the table)
# add the multi actions for a window or multiple windows with w1d2fl r, w2d3fl
# add the keybindings to the prefix commands (with mouse hover read of the window)
param(
    [Parameter(ValueFromRemainingArguments=$true, Position=0)]
    [string[]] $Quick,
    [switch] $Monitors,
    [switch] $Reload,
    [switch] $Activate,
    [Alias('sp')][int] $SplitPercent = 50,
    [Alias('pad','margin')][int] $Spacer = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
function Write-Info($msg) { Write-Host $msg -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host $msg -ForegroundColor Yellow }
function Write-Ok($msg)   { Write-Host $msg -ForegroundColor Green }
# 'Exact' => always match region size; 'Clamp' => shrink only if larger
$FitMode = 'Exact'
<#
    Dynamic load of WindowManager2 with optional force reload.
    When -Reload is used, we compile a new class name (WindowManager2_R<timestamp>)
    and call static members through $WM:: to avoid type caching issues.
#>
# --- Load WindowManager2 (dynamic, supports -Reload) ---
$WMClassName = 'WindowManager2'
if ($Reload) { $WMClassName = 'WindowManager2_R' + [DateTime]::UtcNow.ToString('yyyyMMddHHmmssfff') }
if ($Reload -or -not ($WMClassName -as [type])) {
    Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;
public class $WMClassName {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    public delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);
    [DllImport("user32.dll")] public static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip, MonitorEnumProc lpfnEnum, IntPtr dwData);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    [DllImport("user32.dll")] private static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int left, top, right, bottom; }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct MONITORINFOEX {
        public int cbSize;
        public RECT rcMonitor;  public RECT rcWork;  public uint dwFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string szDevice;
    }
    public class WindowInfo {
        public IntPtr Handle { get; set; }
        public string Title { get; set; }
        public string MonitorName { get; set; }
        public int X { get; set; } public int Y { get; set; }
        public int Width { get; set; } public int Height { get; set; }
    }
    public class MonitorInfo {
        // Name is the Windows device string, e.g. "\\\\.\\DISPLAY2"
        public string Name { get; set; }
        // Parsed display number from device string; matches Control Panel numbering
        public int Number { get; set; }
        // True when monitor is primary (taskbar/main display)
        public bool IsPrimary { get; set; }
        public int X { get; set; } public int Y { get; set; }
        public int Width { get; set; } public int Height { get; set; }
        public int WorkX { get; set; } public int WorkY { get; set; }
        public int WorkWidth { get; set; } public int WorkHeight { get; set; }
    }
    private const int SW_RESTORE = 9;
    private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
    private const uint SWP_NOMOVE = 0x0002; private const uint SWP_NOSIZE = 0x0001;
    public static void EnsureDpiAware() { try { SetProcessDPIAware(); } catch {} }
    public static List<WindowInfo> GetWindows() {
        var windows = new List<WindowInfo>();
        EnumWindows((IntPtr hWnd, IntPtr lParam) => {
            if (!IsWindowVisible(hWnd)) return true;
            int length = GetWindowTextLength(hWnd);
            if (length <= 0) return true;
            var sb = new StringBuilder(length + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            string title = sb.ToString();
            if (string.IsNullOrWhiteSpace(title)) return true;
            IntPtr hMon = MonitorFromWindow(hWnd, 2); // MONITOR_DEFAULTTONEAREST
            var mi = new MONITORINFOEX(); mi.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(MONITORINFOEX));
            if (GetMonitorInfo(hMon, ref mi)) {
                int width = mi.rcMonitor.right - mi.rcMonitor.left;
                int height = mi.rcMonitor.bottom - mi.rcMonitor.top;
                windows.Add(new WindowInfo {
                    Handle = hWnd, Title = title, MonitorName = mi.szDevice,
                    X = mi.rcMonitor.left, Y = mi.rcMonitor.top, Width = width, Height = height
                });
            }
            return true;
        }, IntPtr.Zero);
        return windows;
    }
    public static List<MonitorInfo> GetMonitors() {
        var monitors = new List<MonitorInfo>();
        EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, (IntPtr hMonitor, IntPtr hdc, ref RECT rc, IntPtr data) => {
            var mi = new MONITORINFOEX(); mi.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(MONITORINFOEX));
            if (GetMonitorInfo(hMonitor, ref mi)) {
                int width = mi.rcMonitor.right - mi.rcMonitor.left;
                int height = mi.rcMonitor.bottom - mi.rcMonitor.top;
                int wWidth = mi.rcWork.right - mi.rcWork.left;
                int wHeight = mi.rcWork.bottom - mi.rcWork.top;
                // Parse number from device name ("\\\\.\\DISPLAYN") to match Windows numbering
                int number = 0;
                try {
                    if (!string.IsNullOrEmpty(mi.szDevice)) {
                        var s = mi.szDevice.ToUpperInvariant();
                        int idx = s.LastIndexOf("DISPLAY");
                        if (idx >= 0) {
                            string tail = s.Substring(idx + 7); // after "DISPLAY"
                            int n; if (int.TryParse(tail, out n)) number = n;
                        }
                    }
                } catch { number = 0; }
                bool isPrimary = (mi.dwFlags & 0x00000001) != 0; // MONITORINFOF_PRIMARY
                monitors.Add(new MonitorInfo {
                    Name = mi.szDevice,
                    Number = number,
                    IsPrimary = isPrimary,
                    X = mi.rcMonitor.left, Y = mi.rcMonitor.top,
                    Width = width, Height = height,
                    WorkX = mi.rcWork.left, WorkY = mi.rcWork.top,
                    WorkWidth = wWidth, WorkHeight = wHeight
                });
            }
            return true;
        }, IntPtr.Zero);
        // Sort by Windows display number when available (then by primary as tie-breaker)
        monitors.Sort((a,b) => {
            int na = a.Number <= 0 ? int.MaxValue : a.Number;
            int nb = b.Number <= 0 ? int.MaxValue : b.Number;
            int cmp = na.CompareTo(nb);
            if (cmp != 0) return cmp;
            // Keep primary first if numbers are equal or missing
            if (a.IsPrimary == b.IsPrimary) return 0;
            return a.IsPrimary ? -1 : 1;
        });
        return monitors;
    }
    public static void MoveWindowTo(IntPtr hWnd, int x, int y, int width, int height) {
        MoveWindow(hWnd, x, y, width, height, true);
    }
    public static bool ActivateAllowed = false;
    public static void ActivateWindow(IntPtr hWnd) {
        try { ShowWindow(hWnd, SW_RESTORE); } catch {}
        try { SetForegroundWindow(hWnd); } catch {}
        try {
            SetWindowPos(hWnd, HWND_TOPMOST, 0,0,0,0, SWP_NOMOVE | SWP_NOSIZE);
            SetWindowPos(hWnd, HWND_NOTOPMOST, 0,0,0,0, SWP_NOMOVE | SWP_NOSIZE);
        } catch {}
    }
    public static void MaybeActivateWindow(IntPtr hWnd) {
        if (ActivateAllowed) { ActivateWindow(hWnd); }
    }
}
"@
}
# Resolve type reference for static calls
$WM = $WMClassName -as [type]
# Ensure class has activation gating members; if not, force reload to a fresh class name
$needReload = $false
try { $null = $WM::ActivateAllowed } catch { $needReload = $true }
try { $null = $WM::MaybeActivateWindow([IntPtr]::Zero) } catch { $needReload = $true }
if ($needReload) {
    $WMClassName = 'WindowManager2_R' + [DateTime]::UtcNow.ToString('yyyyMMddHHmmssfff')
    Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;
public class $WMClassName {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    public delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);
    [DllImport("user32.dll")] public static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip, MonitorEnumProc lpfnEnum, IntPtr dwData);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    [DllImport("user32.dll")] private static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int left, top, right, bottom; }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct MONITORINFOEX {
        public int cbSize; public RECT rcMonitor; public RECT rcWork; public uint dwFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string szDevice;
    }
    public class WindowInfo {
        public IntPtr Handle { get; set; }
        public string Title { get; set; }
        public string MonitorName { get; set; }
        public int X { get; set; } public int Y { get; set; }
        public int Width { get; set; } public int Height { get; set; }
    }
    public class MonitorInfo {
        public string Name { get; set; }
        public int Number { get; set; }
        public bool IsPrimary { get; set; }
        public int X { get; set; } public int Y { get; set; }
        public int Width { get; set; } public int Height { get; set; }
        public int WorkX { get; set; } public int WorkY { get; set; }
        public int WorkWidth { get; set; } public int WorkHeight { get; set; }
    }
    private const int SW_RESTORE = 9;
    private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
    private const uint SWP_NOMOVE = 0x0002; private const uint SWP_NOSIZE = 0x0001;
    public static void EnsureDpiAware() { try { SetProcessDPIAware(); } catch {} }
    public static List<WindowInfo> GetWindows() {
        var windows = new List<WindowInfo>();
        EnumWindows((IntPtr hWnd, IntPtr lParam) => {
            if (!IsWindowVisible(hWnd)) return true;
            int length = GetWindowTextLength(hWnd);
            if (length <= 0) return true;
            var sb = new StringBuilder(length + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            string title = sb.ToString();
            if (string.IsNullOrWhiteSpace(title)) return true;
            IntPtr hMon = MonitorFromWindow(hWnd, 2);
            var mi = new MONITORINFOEX(); mi.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(MONITORINFOEX));
            if (GetMonitorInfo(hMon, ref mi)) {
                int width = mi.rcMonitor.right - mi.rcMonitor.left;
                int height = mi.rcMonitor.bottom - mi.rcMonitor.top;
                windows.Add(new WindowInfo { Handle = hWnd, Title = title, MonitorName = mi.szDevice, X = mi.rcMonitor.left, Y = mi.rcMonitor.top, Width = width, Height = height });
            }
            return true;
        }, IntPtr.Zero);
        return windows;
    }
    public static List<MonitorInfo> GetMonitors() {
        var monitors = new List<MonitorInfo>();
        EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, (IntPtr hMonitor, IntPtr hdc, ref RECT rc, IntPtr data) => {
            var mi = new MONITORINFOEX(); mi.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(MONITORINFOEX));
            if (GetMonitorInfo(hMonitor, ref mi)) {
                int width = mi.rcMonitor.right - mi.rcMonitor.left;
                int height = mi.rcMonitor.bottom - mi.rcMonitor.top;
                int wWidth = mi.rcWork.right - mi.rcWork.left;
                int wHeight = mi.rcWork.bottom - mi.rcWork.top;
                int number = 0; try { if (!string.IsNullOrEmpty(mi.szDevice)) { var s = mi.szDevice.ToUpperInvariant(); int idx = s.LastIndexOf("DISPLAY"); if (idx >= 0) { string tail = s.Substring(idx + 7); int n; if (int.TryParse(tail, out n)) number = n; } } } catch {}
                bool isPrimary = (mi.dwFlags & 0x00000001) != 0;
                monitors.Add(new MonitorInfo { Name = mi.szDevice, Number = number, IsPrimary = isPrimary, X = mi.rcMonitor.left, Y = mi.rcMonitor.top, Width = width, Height = height, WorkX = mi.rcWork.left, WorkY = mi.rcWork.top, WorkWidth = wWidth, WorkHeight = wHeight });
            }
            return true;
        }, IntPtr.Zero);
        monitors.Sort((a,b) => { int na = a.Number <= 0 ? int.MaxValue : a.Number; int nb = b.Number <= 0 ? int.MaxValue : b.Number; int cmp = na.CompareTo(nb); if (cmp != 0) return cmp; if (a.IsPrimary == b.IsPrimary) return 0; return a.IsPrimary ? -1 : 1; });
        return monitors;
    }
    public static void MoveWindowTo(IntPtr hWnd, int x, int y, int width, int height) { MoveWindow(hWnd, x, y, width, height, true); }
    public static bool ActivateAllowed = false;
    public static void ActivateWindow(IntPtr hWnd) {
        try { ShowWindow(hWnd, SW_RESTORE); } catch {}
        try { SetForegroundWindow(hWnd); } catch {}
        try { SetWindowPos(hWnd, HWND_TOPMOST, 0,0,0,0, SWP_NOMOVE | SWP_NOSIZE); SetWindowPos(hWnd, HWND_NOTOPMOST, 0,0,0,0, SWP_NOMOVE | SWP_NOSIZE); } catch {}
    }
    public static void MaybeActivateWindow(IntPtr hWnd) { if (ActivateAllowed) { ActivateWindow(hWnd); } }
}
"@
    $WM = $WMClassName -as [type]
}
$WM::ActivateAllowed = $Activate
# --- Helper: window rect (GetWindowRect) ---
if (-not ("Win32RectHelper" -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class Win32RectHelper {
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int left, top, right, bottom; }
    [DllImport("user32.dll", SetLastError = true)] private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    public static int[] GetWindowRectArray(IntPtr hWnd) {
        RECT r; if (GetWindowRect(hWnd, out r)) return new int[] { r.left, r.top, r.right - r.left, r.bottom - r.top };
        return null;
    }
}
'@
}
# --- Helper: window state & detect maximized ---
if (-not ("Win32WindowState" -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class Win32WindowState {
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] private static extern bool IsZoomed(IntPtr hWnd);
    public const int SW_MINIMIZE = 6, SW_MAXIMIZE = 3, SW_RESTORE = 9;
    public static void Minimize(IntPtr hWnd) { ShowWindow(hWnd, SW_MINIMIZE); }
    public static void Maximize(IntPtr hWnd) { ShowWindow(hWnd, SW_MAXIMIZE); }
    public static void Restore (IntPtr hWnd) { ShowWindow(hWnd, SW_RESTORE ); }
    public static bool IsZoomedWindow(IntPtr hWnd) { return IsZoomed(hWnd); }
}
'@
}
# --- Helper: foreground window handle ---
if (-not ("Win32Foreground" -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class Win32Foreground {
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
}
'@
}
# --- Helper: get window title string (simple) ---
if (-not ("Win32TitleHelper" -as [type])) {
    Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class Win32TitleHelper {
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowTextLength(IntPtr hWnd);
    public static string GetTitle(IntPtr hWnd) {
        int len = GetWindowTextLength(hWnd);
        if (len <= 0) return string.Empty;
        var sb = new StringBuilder(len + 1);
        GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }
}
'@
}
# DPI awareness
$WM::EnsureDpiAware()
# --- Enumerate windows & monitors ---
$allWindows  = $WM::GetWindows()
$monitorList = $WM::GetMonitors()
# Normalize monitors to ensure Number/IsPrimary exist even if the type was already loaded
$monitorList = $monitorList | ForEach-Object {
    $num = 0
    if ($_.PSObject.Properties.Match('Number').Count -gt 0 -and $_.Number) {
        $num = [int]$_.Number
    } else {
        try {
            if ($_.PSObject.Properties.Match('Name').Count -gt 0 -and [string]::IsNullOrEmpty($_.Name) -eq $false) {
                $s = $_.Name.ToUpperInvariant()
                $idx = $s.LastIndexOf('DISPLAY')
                if ($idx -ge 0) {
                    $tail = $s.Substring($idx + 7)
                    [int]::TryParse($tail, [ref]$num) | Out-Null
                }
            }
        } catch { $num = 0 }
    }
    $isPrimary = $false
    if ($_.PSObject.Properties.Match('IsPrimary').Count -gt 0) { $isPrimary = [bool]$_.IsPrimary }
    [pscustomobject]@{
        Name = $_.Name
        Number = $num
        IsPrimary = $isPrimary
        X = $_.X; Y = $_.Y
        Width = $_.Width; Height = $_.Height
        WorkX = $_.WorkX; WorkY = $_.WorkY
        WorkWidth = $_.WorkWidth; WorkHeight = $_.WorkHeight
    }
} | Sort-Object -Property @{ Expression = { if ($_.Number -gt 0) { $_.Number } else { [int]::MaxValue } } }, @{ Expression = { -not $_.IsPrimary } }
if ($Monitors) {
    Write-Info "`nAvailable Monitors (Windows numbering):"
    for ($j = 0; $j -lt $monitorList.Count; $j++) {
        $m = $monitorList[$j]
        $num = if ($m.Number -gt 0) { $m.Number } else { ($j+1) }
        $primary = if ($m.IsPrimary) { ' (Primary)' } else { '' }
        "{0,2}. Display {1}{10} | Device: {2} | Resolution: {3}x{4} @ ({5},{6}) | WorkArea: {7}x{8} @ ({9},{11})" -f `
            ($j+1), $num, $m.Name, $m.Width, $m.Height, $m.X, $m.Y, $m.WorkWidth, $m.WorkHeight, $m.WorkX, $primary, $m.WorkY | Write-Host
    }
    return
}
if (-not $allWindows -or $allWindows.Count -eq 0) { Write-Warn "No visible titled windows found."; return }
if (-not $monitorList -or $monitorList.Count -eq 0)    { Write-Warn "No monitors detected."; return }
# =======================
# History (Undo / Redo)
# =======================
if (-not (Test-Path variable:global:UndoStack)) { $global:UndoStack = New-Object System.Collections.ArrayList }
if (-not (Test-Path variable:global:RedoStack)) { $global:RedoStack = New-Object System.Collections.ArrayList }
function Get-Rect($handle) {
    $r = [Win32RectHelper]::GetWindowRectArray($handle)
    if ($null -eq $r) { return $null }
    [pscustomobject]@{ X=$r[0]; Y=$r[1]; W=$r[2]; H=$r[3] }
}
function Push-Undo([System.IntPtr]$handle, $fromRect, $toRect) {
    [void]$global:UndoStack.Add([pscustomobject]@{ Handle = $handle; From = $fromRect; To = $toRect; Timestamp = Get-Date })
    $null = $global:RedoStack.Clear()
}
function Do-Undo() {
    if ($global:UndoStack.Count -eq 0) { Write-Warn "Nothing to undo."; return }
    $entry = $global:UndoStack[$global:UndoStack.Count-1]
    [void]$global:UndoStack.RemoveAt($global:UndoStack.Count-1)
    $cur = Get-Rect $entry.Handle
    if ($cur) {
        $WM::MoveWindowTo($entry.Handle, $entry.From.X, $entry.From.Y, $entry.From.W, $entry.From.H)
        $WM::MaybeActivateWindow($entry.Handle)
        [void]$global:RedoStack.Add([pscustomobject]@{ Handle = $entry.Handle; From = $cur; To = $entry.From; Timestamp = Get-Date })
        Write-Ok ("Undo → ({0},{1},{2}x{3})" -f $entry.From.X, $entry.From.Y, $entry.From.W, $entry.From.H)
    }
}
function Do-Redo() {
    if ($global:RedoStack.Count -eq 0) { Write-Warn "Nothing to redo."; return }
    $entry = $global:RedoStack[$global:RedoStack.Count-1]
    [void]$global:RedoStack.RemoveAt($global:RedoStack.Count-1)
    $cur = Get-Rect $entry.Handle
    if ($cur) {
        $WM::MoveWindowTo($entry.Handle, $entry.To.X, $entry.To.Y, $entry.To.W, $entry.To.H)
        $WM::MaybeActivateWindow($entry.Handle)
        [void]$global:UndoStack.Add([pscustomobject]@{ Handle = $entry.Handle; From = $cur; To = $entry.To; Timestamp = Get-Date })
        Write-Ok ("Redo → ({0},{1},{2}x{3})" -f $entry.To.X, $entry.To.Y, $entry.To.W, $entry.To.H)
    }
}
# =======================
# Helpers (geometry & current window)
# =======================
function New-Rect([int]$x,[int]$y,[int]$w,[int]$h) { [pscustomobject]@{ X=$x; Y=$y; W=$w; H=$h } }
function Split-Rect($rect, [char]$dir) {
    $p = [math]::Max(1, [math]::Min(99, [int]$SplitPercent))
    switch ($dir.ToString().ToLowerInvariant()) {
        'l' {
            $w = [math]::Floor($rect.W * $p / 100)
            return New-Rect $rect.X $rect.Y $w $rect.H
        }
        'r' {
            $w = [math]::Floor($rect.W * $p / 100)
            $x = $rect.X + ($rect.W - $w)
            return New-Rect $x $rect.Y $w $rect.H
        }
        't' {
            $h = [math]::Floor($rect.H * $p / 100)
            return New-Rect $rect.X $rect.Y $rect.W $h
        }
        'b' {
            $h = [math]::Floor($rect.H * $p / 100)
            $y = $rect.Y + ($rect.H - $h)
            return New-Rect $rect.X $y $rect.W $h
        }
        'f' { return New-Rect $rect.X $rect.Y $rect.W $rect.H }
        default { throw "Invalid direction '$dir'. Use only l/r/t/b/f." }
    }
}
function Apply-Directions($baseRect, [string]$dirs) {
    if ([string]::IsNullOrWhiteSpace($dirs)) { return $baseRect }
    $rect = $baseRect
    foreach ($c in $dirs.ToCharArray()) { $rect = Split-Rect $rect $c }
    return $rect
}
function Apply-Grid($baseRect, [int]$cols, [int]$rows, [int]$rowIdx, [int]$colIdx) {
    if ($cols -lt 1 -or $rows -lt 1) { throw "Grid must be at least 1x1." }
    if ($rowIdx -lt 1 -or $rowIdx -gt $rows) { throw "Row must be 1..$rows." }
    if ($colIdx -lt 1 -or $colIdx -gt $cols) { throw "Column must be 1..$cols." }
    $cellW = [math]::Floor($baseRect.W / $cols); $cellH = [math]::Floor($baseRect.H / $rows)
    $x = $baseRect.X + $cellW * ($colIdx - 1); $y = $baseRect.Y + $cellH * ($rowIdx - 1)
    $w = $cellW; if ($colIdx -eq $cols) { $w = $baseRect.W - $cellW * ($cols - 1) }
    $h = $cellH; if ($rowIdx -eq $rows) { $h = $baseRect.H - $cellH * ($rows - 1) }
    return New-Rect $x $y $w $h
}
$QuadrantDirs = @{ 'q1'='tl'; 'q2'='tr'; 'q3'='bl'; 'q4'='br' }
function Apply-Spacer($rect, [int]$pad) {
    if ($pad -le 0) { return $rect }
    $x = $rect.X + $pad
    $y = $rect.Y + $pad
    $w = $rect.W - 2*$pad
    $h = $rect.H - 2*$pad
    if ($w -lt 1) { $w = 1 }
    if ($h -lt 1) { $h = 1 }
    return New-Rect $x $y $w $h
}
function Fit-IntoRegion($currentRect, $regionRect, [string]$mode) {
    if ($mode -eq 'Clamp' -and $currentRect) {
        $newW = [math]::Min($currentRect.W, $regionRect.W)
        $newH = [math]::Min($currentRect.H, $regionRect.H)
        return New-Rect $regionRect.X $regionRect.Y $newW $newH
    } else { return $regionRect }
}
function Get-CurrentWindowInfo {
    $h = [Win32Foreground]::GetForegroundWindow()
    if ($h -eq [IntPtr]::Zero) { return $null }
    $r = Get-Rect $h
    if (-not $r) { return $null }
    $cx = $r.X + [math]::Floor($r.W / 2)
    $cy = $r.Y + [math]::Floor($r.H / 2)
    $mon = $monitorList | Where-Object {
        $cx -ge $_.X -and $cx -le ($_.X + $_.Width) -and
        $cy -ge $_.Y -and $cy -le ($_.Y + $_.Height)
    } | Select-Object -First 1
    if (-not $mon) { $mon = $monitorList[0] }
    $title = [Win32TitleHelper]::GetTitle($h)
    [pscustomobject]@{
        Handle = $h
        Title = $title
        MonitorName = $mon.Name
    }
}
# =======================
# Move helpers with history (restores maximized first)
# =======================
function Move-WindowWithHistory($winObj, $targetRegionRect) {
    try {
        if ([Win32WindowState]::IsZoomedWindow($winObj.Handle)) {
            [Win32WindowState]::Restore($winObj.Handle)
            Start-Sleep -Milliseconds 60
        }
    } catch {}
    $fromRect = Get-Rect $winObj.Handle
    $targetRect = Fit-IntoRegion -currentRect $fromRect -regionRect $targetRegionRect -mode $FitMode
    # Apply optional spacer/margin
    $targetRect = Apply-Spacer -rect $targetRect -pad $Spacer
    $WM::MoveWindowTo($winObj.Handle, $targetRect.X, $targetRect.Y, $targetRect.W, $targetRect.H)
    $WM::MaybeActivateWindow($winObj.Handle)
    $toRect = Get-Rect $winObj.Handle
    if ($fromRect -and $toRect) { Push-Undo -handle $winObj.Handle -fromRect $fromRect -toRect $toRect }
}
function Set-WindowStateWithHistory($winObj, [string]$action, $targetMonitorOrNull) {
    $fromRect = Get-Rect $winObj.Handle
    if ($action -eq 'maximize') {
        if ($null -ne $targetMonitorOrNull) {
            $base = New-Rect $targetMonitorOrNull.WorkX $targetMonitorOrNull.WorkY $targetMonitorOrNull.WorkWidth $targetMonitorOrNull.WorkHeight
            $WM::MoveWindowTo($winObj.Handle, $base.X, $base.Y, $base.W, $base.H)
        }
        [Win32WindowState]::Maximize($winObj.Handle)
        $WM::MaybeActivateWindow($winObj.Handle)
    } elseif ($action -eq 'minimize') {
        [Win32WindowState]::Minimize($winObj.Handle)
    } elseif ($action -eq 'restore') {
        [Win32WindowState]::Restore($winObj.Handle)
        $WM::MaybeActivateWindow($winObj.Handle)
    }
    $toRect = Get-Rect $winObj.Handle
    if ($fromRect -and $toRect) { Push-Undo -handle $winObj.Handle -fromRect $fromRect -toRect $toRect }
}
# =======================
# Quick Command Parser (d = down alias; current-window fallback)
# =======================
function Parse-And-ExecuteQuick(
    [string]$qcInput,
    [System.Collections.ArrayList]$windowsList,
    $defaultSingleWindow
) {
    $qcRaw = ($qcInput.Trim())
    $qcLower = $qcRaw.ToLowerInvariant()
    # Undo/redo
    if ($qcLower -eq 'undo') { Do-Undo; return $true }
    if ($qcLower -eq 'redo') { Do-Redo; return $true }
    # Tokens: wN, dM or sdM (any order). NOTE: 'd' w/o digits is 'down' handled later.
    $wMatch = [regex]::Match($qcLower, 'w(?<w>\d+)')
    $dMatches = [regex]::Matches($qcLower, '(?:sd|d)(?<d>\d+)')
    $wIdx = $null; if ($wMatch.Success) { $wIdx = [int]$wMatch.Groups['w'].Value }
    $dIdx = $null; if ($dMatches.Count -gt 0) { $dIdx = [int]$dMatches[$dMatches.Count-1].Groups['d'].Value }
    # Grid & Quadrant
    $gridMatch = [regex]::Match($qcLower, '(?<cols>\d+)x(?<rows>\d+):r(?<r>\d+)c(?<c>\d+)')
    $quadMatch = [regex]::Match($qcLower, 'q[1-4]')
    # Strip tokens: display (dN|sdN), window (wN), grid, quad → remaining is dirs/actions
    $rest = [regex]::Replace($qcRaw, '(?i)(?:sd|d)\d+|w\d+|\d+x\d+:r\d+c\d+|q[1-4]', '')
    $restTrim = $rest.Trim()
    # Actions m/M/min/max/minimize/maximize
    $action = $null
    if ($restTrim -match '(?i)\bmax(imize)?\b' -or $restTrim.Contains('M')) { $action = 'maximize' }
    elseif ($restTrim -match '(?i)\bmin(imize)?\b' -or $restTrim -match '(^|[^a-zA-Z])m([^a-zA-Z]|$)') { $action = 'minimize' }
    # Activation token: accept word 'activate' or any bare 'a' left after stripping action tokens (supports suffix like '...a')
    $restSansActions = [regex]::Replace($restTrim, '(?i)max(imize)?|min(imize)?|M', '')
    $activateToken = ($restSansActions -match '(?i)\bactivate\b' -or $restSansActions.ToLowerInvariant().Contains('a'))
    # Directions from remaining letters: allow l/r/t/b/f and also 'd' (down=bottom)
    $dirsPart = $restTrim
    $dirsPart = [regex]::Replace($dirsPart, '(?i)max(imize)?|min(imize)?|M|\bactivate\b', '')
    $dirsPart = $dirsPart.ToLowerInvariant()
    $dirs = ($dirsPart -replace '[^dlrtbf]', '')
    if ($dirs) { $dirs = $dirs -replace 'd','b' }  # map 'd' => 'b' (down=bottom)
    # --- Resolve window ---
    $selWin = $null
    if ($wIdx -ne $null) {
        if ($wIdx -lt 1 -or $wIdx -gt $windowsList.Count) { Write-Warn "Window index out of range (1..$($windowsList.Count))."; return $false }
        $selWin = $windowsList[$wIdx-1]
    } elseif ($defaultSingleWindow) {
        $selWin = $defaultSingleWindow
    } else {
        # Fallback: current foreground window
        $selWin = Get-CurrentWindowInfo
        if (-not $selWin) { Write-Warn "Could not resolve current foreground window."; return $false }
    }
    # --- Resolve TARGET monitor FIRST (bug fix) ---
    $tMon = $null
    if ($dIdx -ne $null) {
        if ($dIdx -lt 1 -or $dIdx -gt $monitorList.Count) { Write-Warn "Display index out of range (1..$($monitorList.Count))."; return $false }
        $tMon = $monitorList[$dIdx-1]
    } else {
        $tMon = $monitorList | Where-Object { $_.Name -eq $selWin.MonitorName } | Select-Object -First 1
        if (-not $tMon) { $tMon = $monitorList[0] }
    }
    # --- Early exit for actions (min/max) ---
    if ($action) {
        $prevActivate = $WM::ActivateAllowed
        # Allow activation if -Activate parameter or 'a' token is used
        $WM::ActivateAllowed = ($Activate -or $activateToken)
        Set-WindowStateWithHistory -winObj $selWin -action $action -targetMonitorOrNull $(if ($action -eq 'maximize') { $tMon } else { $null })
        $WM::ActivateAllowed = $prevActivate
        Write-Ok ("{0} '{1}' on {2}" -f ($action.Substring(0,1).ToUpper()+$action.Substring(1)), $selWin.Title, $tMon.Name)
        return $true
    }
    # --- Build base FROM TARGET monitor (not source) ---
    $base = New-Rect $tMon.WorkX $tMon.WorkY $tMon.WorkWidth $tMon.WorkHeight
    # --- Choose target region on that base ---
    $targetRegion = $null
    if ($gridMatch.Success) {
        $cols = [int]$gridMatch.Groups['cols'].Value
        $rows = [int]$gridMatch.Groups['rows'].Value
        $row  = [int]$gridMatch.Groups['r'].Value
        $col  = [int]$gridMatch.Groups['c'].Value
        $targetRegion = Apply-Grid $base $cols $rows $row $col
    } elseif ($quadMatch.Success) {
        $qdDirs = $QuadrantDirs[$quadMatch.Value]
        $targetRegion = Apply-Directions $base $qdDirs
    } else {
        $targetRegion = Apply-Directions $base $dirs
    }
    # --- Move (restore if maximized), activate, history ---
    $prevActivate2 = $WM::ActivateAllowed
    $WM::ActivateAllowed = ($Activate -or $activateToken)
    Move-WindowWithHistory -winObj $selWin -targetRegionRect $targetRegion
    $WM::ActivateAllowed = $prevActivate2
    # --- Output ---
    $desc = $null
    if     ($gridMatch.Success)                  { $desc = $gridMatch.Value }
    elseif ($quadMatch.Success)                  { $desc = $quadMatch.Value }
    elseif ([string]::IsNullOrWhiteSpace($dirs)) { $desc = '(none)' } else { $desc = $dirs }
    Write-Ok ("Moved '{0}' to {1} on {2} → ({3},{4},{5}x{6})." -f `
        $selWin.Title, $desc, $tMon.Name, $targetRegion.X, $targetRegion.Y, $targetRegion.W, $targetRegion.H)
    return $true
}
# =======================
# CLI Modes (direct quick / search-first / interactive)
# =======================
function Test-IsQuickCommand([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return $false }
    $t = $s.Trim()
    # wN | dN/sdN | grid | quadrant | pure dirs/actions (including 'd' down, m/M)
    return ($t -match '(?i)w\d+' -or $t -match '(?i)(?:sd|d)\d+' -or $t -match '^\d+x\d+:r\d+c\d+$' -or $t -match '(?i)^q[1-4]$' -or $t -match '^[dlrtbf]+$' -or $t -match '(?i)\b(min|max|minimize|maximize)\b' -or $t -match '\b[Mm]\b')
}
$filteredWindows = $allWindows
$cli1 = $null; $cli2 = $null
$__src = @()
if ($null -ne $Quick -and @($Quick).Count -gt 0) { $__src = @($Quick) } else { $__src = @($args) }
if ($__src.Count -ge 1) { $cli1 = [string]$__src[0] }
if ($__src.Count -ge 2) { $cli2 = [string]$__src[1] }
# Ignore leading flag-only inputs (e.g., -Spacer 0) so they don't become search terms
if ($cli1 -and $cli1.StartsWith('-')) { $cli1 = $null; $cli2 = $null }
# Case 1: <search> <quick>
    if ($cli1 -and $cli2) {
        $pattern = [Regex]::Escape($cli1)
        $filteredWindows = @($allWindows | Where-Object { $_.Title -match $pattern })
        if (-not $filteredWindows -or $filteredWindows.Count -eq 0) { Write-Warn "No windows match '$cli1'."; return }
        if ($filteredWindows.Count -eq 1) {
        $executed = Parse-And-ExecuteQuick -qcInput $cli2 `
            -windowsList ([System.Collections.ArrayList]$filteredWindows) `
            -defaultSingleWindow $filteredWindows[0]
            if ($executed) { return }
        }
    }
# Case 2: 1 arg that looks like a quick command → direct execute (no search)
elseif ($cli1 -and (Test-IsQuickCommand $cli1)) {
    $fullArray = [System.Collections.ArrayList]@($allWindows)
    $executed = Parse-And-ExecuteQuick -qcInput $cli1 -windowsList $fullArray -defaultSingleWindow $null
    if ($executed) { return }
}
# Case 3: 1 arg is a search (not a flag)
elseif ($cli1 -and -not $cli1.StartsWith('-')) {
    $pattern = [Regex]::Escape($cli1)
    $filteredWindows = @($allWindows | Where-Object { $_.Title -match $pattern })
    if (-not $filteredWindows -or $filteredWindows.Count -eq 0) { Write-Warn "No windows match '$cli1'."; return }
    Write-Info "`nFiltered Windows for '$cli1':"
} else {
    Write-Info "`nAvailable Windows:"
    $filteredWindows = @($filteredWindows)
}
# Print lists for interactivity
for ($i = 0; $i -lt $filteredWindows.Count; $i++) {
    $w = $filteredWindows[$i]
    "{0,2}. {1} | Current Monitor: {2}" -f ($i+1), $w.Title, $w.MonitorName | Write-Host
}
Write-Host "Options: l/r/t/b/d/f (d=down) | Quadrants: q1..q4 | Grid: NxM:rRcC | Actions: m=minimize, M=maximize, a=activate" -ForegroundColor DarkCyan
Write-Info "`nAvailable Monitors (Windows numbering):"
for ($j = 0; $j -lt $monitorList.Count; $j++) {
    $m = $monitorList[$j]
    $num = if ($m.Number -gt 0) { $m.Number } else { ($j+1) }
    $primary = if ($m.IsPrimary) { ' (Primary)' } else { '' }
    "{0,2}. Display {1}{10} | Device: {2} | Resolution: {3}x{4} @ ({5},{6}) | WorkArea: {7}x{8} @ ({9},{11})" -f `
        ($j+1), $num, $m.Name, $m.Width, $m.Height, $m.X, $m.Y, $m.WorkWidth, $m.WorkHeight, $m.WorkX, $primary, $m.WorkY | Write-Host
}
# =======================
# Interactive quick prompt (optional)
# =======================
$qc = Read-Host "`nEnter quick command (e.g., w5tl, d1f, ddl, 3x3:r2c1, q2, undo, redo, w1m, w1M; add 'a' to activate). Global: -SplitPercent $SplitPercent, -Spacer $Spacer. Press Enter for menus"
if (-not [string]::IsNullOrWhiteSpace($qc)) {
    $defaultSingle = $null
    if ($filteredWindows.Count -eq 1) { $defaultSingle = $filteredWindows[0] }
    $executed = Parse-And-ExecuteQuick -qcInput $qc `
        -windowsList ([System.Collections.ArrayList]$filteredWindows) `
        -defaultSingleWindow $defaultSingle
    if ($executed) { return }
}
# =======================
# MENU MODE (fallback)
# =======================
# 1) Pick a window
$selected = $null
while (-not $selected) {
    $choice = Read-Host "`nSelect window number (or 'undo'/'redo', or Enter=current window)"
    if ([string]::IsNullOrWhiteSpace($choice)) {
        $selected = Get-CurrentWindowInfo
        if (-not $selected) { Write-Warn "Cancelled."; return }
        break
    }
    $cLower = $choice.ToLowerInvariant()
    if ($cLower -eq 'undo') { Do-Undo; continue }
    if ($cLower -eq 'redo') { Do-Redo; continue }
    $tmp = 0
    if ([int]::TryParse($choice, [ref]$tmp)) {
        $idx = [int]$tmp
        if ($idx -ge 1 -and $idx -le $filteredWindows.Count) { $selected = $filteredWindows[$idx - 1] }
    }
    if (-not $selected) { Write-Warn "Invalid selection, try again." }
}
# 2) Pick a monitor
$targetMonitor = $null
while (-not $targetMonitor) {
$monitorChoice = Read-Host "Select target monitor number (Enter = current monitor)"
    if ([string]::IsNullOrWhiteSpace($monitorChoice)) {
        $targetMonitor = $monitorList | Where-Object { $_.Name -eq $selected.MonitorName } | Select-Object -First 1
        if (-not $targetMonitor) { $targetMonitor = $monitorList[0] }
        break
    }
    $tmp2 = 0
    if ([int]::TryParse($monitorChoice, [ref]$tmp2)) {
        $mid = [int]$tmp2
        if ($mid -ge 1 -and $mid -le $monitorList.Count) { $targetMonitor = $monitorList[$mid - 1] }
    }
    if (-not $targetMonitor) { Write-Warn "Invalid selection, try again." }
}
# 3) Placement / Action
Write-Host "Options: l/r/t/b/d/f | Quadrants: q1..q4 | Grid: NxM:rRcC | Actions: m=minimize, M=maximize, a=activate | Global: -SplitPercent $SplitPercent, -Spacer $Spacer" -ForegroundColor DarkCyan
$dirs = $null; $gridSpec = $null; $quadSpec = $null; $action = $null
while (-not $dirs -and -not $gridSpec -and -not $quadSpec -and -not $action) {
    $inp = Read-Host "Enter dirs (tl/rrb/f/ddl), quadrant (q1..q4), grid (3x3:r2c1), or action (m/M). Enter to cancel"
    if ([string]::IsNullOrWhiteSpace($inp)) { Write-Warn "Cancelled."; return }
    $inp = $inp.Trim()
    if ($inp -match '^\d+x\d+:r\d+c\d+$') { $gridSpec = $inp.ToLowerInvariant(); break }
    if ($inp -match '^(?i)q[1-4]$')       { $quadSpec = $inp.ToLowerInvariant(); break }
    if ($inp -match '^[dlrtbf]+$')        { $dirs = $inp.ToLowerInvariant(); break }
    if ($inp -match '^(?i)(m|min|minimize)$') { $action = 'minimize'; break }
    if ($inp -match '^(?i)(M|max|maximize)$') { $action = 'maximize'; break }
    # words
    $map = @{ left='l'; right='r'; top='t'; bottom='b'; down='d'; full='f'; fullscreen='f'; max='M'; maximize='M'; min='m'; minimize='m' }
    if ($map.ContainsKey($inp.ToLowerInvariant())) {
        $val = $map[$inp.ToLowerInvariant()]
        if ($val -eq 'M'){ $action='maximize' } elseif ($val -eq 'm'){ $action='minimize' } else { $dirs=$val }
        break
    }
    Write-Warn "Invalid input. Examples: tl, rrb, f, q2, 3x3:r2c1, m, M, ddl"
}
    if ($action) {
        $prevActivate = $WM::ActivateAllowed
        $WM::ActivateAllowed = $Activate
        Set-WindowStateWithHistory -winObj $selected -action $action -targetMonitorOrNull $(if ($action -eq 'maximize') { $targetMonitor } else { $null })
        $WM::ActivateAllowed = $prevActivate
        Write-Ok ("{0} '{1}' on {2}" -f ($action.Substring(0,1).ToUpper()+$action.Substring(1)), $selected.Title, $targetMonitor.Name)
        return
    }
$baseRect = New-Rect $targetMonitor.WorkX $targetMonitor.WorkY $targetMonitor.WorkWidth $targetMonitor.WorkHeight
if     ($gridSpec) {
    $m = [regex]::Match($gridSpec, '^(?<cols>\d+)x(?<rows>\d+):r(?<r>\d+)c(?<c>\d+)$')
    $targetRegion = Apply-Grid $baseRect ([int]$m.Groups['cols'].Value) ([int]$m.Groups['rows'].Value) ([int]$m.Groups['r'].Value) ([int]$m.Groups['c'].Value)
} elseif ($quadSpec) {
    $dirs = $QuadrantDirs[$quadSpec]
    $targetRegion = Apply-Directions $baseRect $dirs
} else {
    if ($dirs) { $dirs = $dirs -replace 'd','b' } # down -> bottom
    $targetRegion = Apply-Directions $baseRect $dirs
}
    $prevActivate2 = $WM::ActivateAllowed
    $WM::ActivateAllowed = $Activate
    Move-WindowWithHistory -winObj $selected -targetRegionRect $targetRegion
    $WM::ActivateAllowed = $prevActivate2
Write-Ok ("Moved '{0}' on {1} → ({2},{3},{4}x{5})." -f `
    $selected.Title, $targetMonitor.Name, $targetRegion.X, $targetRegion.Y, $targetRegion.W, $targetRegion.H)
``

