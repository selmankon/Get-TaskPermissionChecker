param(
    [switch]$All,
    [switch]$cls,
    [switch]$Debug
)

function Write-DebugInfo {
    param(
        [string]$message
    )
    if ($Debug.IsPresent) {
        Write-Host $message -ForegroundColor Magenta
    }
}

if ($cls.IsPresent) {
    Clear-Host
}

Write-Host " "
Write-Host "Scheduled Task Comprehensive Analysis" -ForegroundColor Cyan
Write-Host "Waiting for schtasks to complete..." -ForegroundColor Cyan
Write-Host "Use -clearScreen switch to clear the screen." -ForegroundColor Yellow
Write-Host " "

$tasks = schtasks /query /fo CSV /v | ConvertFrom-Csv

$currentUsername = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$userPrincipal = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
$userGroups = $userPrincipal.Identity.Groups | ForEach-Object { $_.Translate([System.Security.Principal.NTAccount]) }

$canModifyCount = 0
$cannotModifyCount = 0

foreach ($task in $tasks) {
    $taskToRun = $task.'Task To Run'
    $taskName = $task.'TaskName'
    if (-not [string]::IsNullOrWhiteSpace($taskToRun) -and $taskToRun -notmatch "COM handler" -and $taskName -notmatch "TaskName") {
        $expandedPath = [Environment]::ExpandEnvironmentVariables($taskToRun).Trim()

        if ($expandedPath -match '^(.*?\.exe)') {
            $exePath = $matches[1]

            if ($exePath.StartsWith('"') -and $exePath.EndsWith('"')) {
                $exePath = $exePath.Substring(1, $exePath.Length - 2)
            }

            if (Test-Path $exePath) {
                $permissions = icacls "`"$exePath`"" | Out-String
                $canModify = $false

                foreach ($group in $userGroups) {
                    if ($permissions -match [regex]::Escape($group.Value) + '.*(F|M|W)') {
                        $canModify = $true
                        break
                    }
                }

                if (-not $canModify -and $permissions -match [regex]::Escape($currentUsername) + '.*(F|M|W)') {
                    $canModify = $true
                }

                $actionColor = if ($canModify) {"Red"} else {"black"}
                $additionalMessage = if (-not ($exePath -imatch "microsoft|windows")) {" (also, there is no Windows path)"} else {""}

                if ($canModify -or $All.IsPresent) {
                    Write-Host "`nTask Information for: $($task.TaskName)" -ForegroundColor Green
                    Write-Host "Current User: $currentUsername can action this file.$additionalMessage" -ForegroundColor $actionColor
                    foreach ($property in $task.PSObject.Properties) {
                        if (-not [string]::IsNullOrWhiteSpace($property.Value)) {
                            Write-Host "$($property.Name): $($property.Value)" -ForegroundColor Yellow
                        }
                    }
                    if ($canModify) { 
                        $canModifyCount++ 
                    } else { 
                        $cannotModifyCount++ 
                    }
                }
            } else {
                $cannotModifyCount++
                Write-DebugInfo "Executable file not found or inaccessible for task: $($task.TaskName), Path: $exePath"
                if ($All.IsPresent) {
                    Write-Host "`nTask Information for: $($task.TaskName)" -ForegroundColor Red
                    Write-Host "Executable file not found: `"$exePath`"" -ForegroundColor Red
                }
            }
        } else {
            $cannotModifyCount++
            Write-DebugInfo "No executable path matched or task does not specify a valid '.exe' for: $($task.TaskName)"
            if ($All.IsPresent) {
                Write-Host "`nTask Information for: $($task.TaskName)" -ForegroundColor Red
                Write-Host "Task does not specify a valid '.exe' file or path" -ForegroundColor Red
            }
        }
    }
}

Write-Host " "
Write-Host "`nTotal tasks user can action: $canModifyCount" -ForegroundColor Green
Write-Host "Total tasks user cannot action (or missing files): $cannotModifyCount" -ForegroundColor Red
if (-not $All.IsPresent) {
    Write-Host "Use -All switch to see all tasks including those you cannot action or missing executable files." -ForegroundColor Yellow
}
