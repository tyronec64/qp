# qpw — Quick Path PowerShell Window Control

`qpw` is a PowerShell command for quickly selecting, focusing, and snapping windows using short tokens and TAB-completion.

The project consists of:

1. **`qpw.ps1`** — the window control command
2. **Profile integration** — wrapper + argument completion
3. **Installer (`Install-Qpw.ps1`)** — interactive setup for PS5.1 and PS7+

---

## Requirements

- Windows
- PowerShell:
  - **Windows PowerShell 5.1**
  - **PowerShell 7+**
- No mandatory external modules

---

## Repository Layout

```text
qpw.ps1
Install-Qpw.ps1
README.md
CHANGELOG.md

## After instalaltion

```
Documents\
 ├─ PowerShell\                  # PowerShell 7+
 │   ├─ qpw.ps1
 │   ├─ qpw-Microsoft.PowerShell_profile.ps1
 │   └─ Microsoft.PowerShell_profile.ps1
 └─ WindowsPowerShell\           # Windows PowerShell 5.1
     ├─ qpw.ps1
     ├─ qpw-Microsoft.PowerShell_profile.ps1
     └─ Microsoft.PowerShell_profile.ps1
```

## Installation

```
git clone <repo-url>
cd qpw
```

```
.\Install-Qpw.ps1
```

## Troubleshooting

### Execution policy


If execution policy blocks scripts:

```
powershell -ExecutionPolicy Bypass -File .\Install-Qpw.ps1
# PowerShell 7+
pwsh -ExecutionPolicy Bypass -File .\Install-Qpw.ps1
```

if you see "running scripts is disabled on this system"
Fix once (recommended):

```
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

if files were down;loaded/onedrive-marked:

```
Unblock-File ~/Documents/PowerShell/qpw.ps1
Unblock-File ~/Documents/WindowsPowerShell/qpw.ps1
```

### Profile path

there is a difference in ps5 and ps7+ in profile paths

#### psh 7+
```
$env:USERPROFILE\Documents\PowerShell\Microsoft.PowerShell_profile.ps1
```

#### psh 5.1

```
$env:USERPROFILE\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1
``` 

### Profile path refresh
reload the current profile after installation.

```
. $profile
```

Or do it with the seperate supplied snippets.

#### ps7+
```
. "$env:USERPROFILE\Documents\PowerShell\qpw-Microsoft.PowerShell_profile.ps1"
```

#### ps5+
```
. "$env:USERPROFILE\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1"

```


## Usage

```
qpw
qpw chrome
qpw "Visual Studio"
qpw w2 snap r
qpw snap q1
```

### Token Examples

| Token          | Meaning         |
| -------------- | --------------- |
| `w1`, `w2`     | Window index    |
| `d1`, `d2`     | Monitor index   |
| `l r u d`      | Snap directions |
| `q1..q4`       | Quadrants       |
| `snap`         | Snap action     |
| `undo`, `redo` | Actions         |

### Tab Completion

```
qpw <TAB>
qpw chr<TAB>
qpw snap <TAB>
qpw w<TAB>
```
Works in:
  Windows PowerShell 5.1
  PowerShell 7+
No global TAB behavior is changed.

