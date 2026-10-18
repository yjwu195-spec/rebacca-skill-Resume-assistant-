[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$ProjectRoot,
    [Parameter(Mandatory=$true)][string]$ResumeRoot,
    [Parameter(Mandatory=$true)][string]$TrackerPath,
    [string]$SkillsRoot = $(if ($env:CODEX_HOME) { Join-Path $env:CODEX_HOME 'skills' } else { Join-Path $env:USERPROFILE '.codex\skills' })
)
$ErrorActionPreference = 'Stop'
$resolvedProject = (Resolve-Path -LiteralPath $ProjectRoot).Path
$resolvedResume = (Resolve-Path -LiteralPath $ResumeRoot).Path
$resolvedTracker = (Resolve-Path -LiteralPath $TrackerPath).Path
if (-not (Test-Path -LiteralPath $resolvedProject -PathType Container)) { throw 'ProjectRoot must be a directory.' }
if (-not (Test-Path -LiteralPath $resolvedResume -PathType Container)) { throw 'ResumeRoot must be a directory.' }
if (-not (Test-Path -LiteralPath $resolvedTracker -PathType Leaf)) { throw 'TrackerPath must be an existing workbook.' }
if ([IO.Path]::GetExtension($resolvedTracker) -ne '.xlsx') { throw 'TrackerPath must point to an xlsx workbook.' }
$sourceSkill = Join-Path $PSScriptRoot 'skills\rebacca'
$destinationSkill = Join-Path $SkillsRoot 'rebacca'
if (Test-Path -LiteralPath $destinationSkill) { throw 'rebacca already exists. Inspect the existing skill before replacing it.' }
$null = New-Item -ItemType Directory -Path $SkillsRoot -Force
Copy-Item -LiteralPath $sourceSkill -Destination $destinationSkill -Recurse
$settings = [ordered]@{ project_root=$resolvedProject; resume_root=$resolvedResume; tracker_path=$resolvedTracker }
$utf8 = New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText((Join-Path $destinationSkill 'local-settings.json'), ($settings | ConvertTo-Json), $utf8)
$router = @"

<!-- rebacca-explicit-entry -->
仅当用户有意输入 `$Rebacca 或 `$rebacca，或说“调用求职投递助手 skill”，读取技能文件：$destinationSkill\SKILL.md。
单独 JD、岗位问题、写邮件不启动完整流程。每次只处理当次岗位，状态和 sheet 回复可延续本次任务，下一份 JD 需重新调用。
<!-- /rebacca-explicit-entry -->
"@
$agentsPath = Join-Path $resolvedProject 'AGENTS.md'
if (Test-Path -LiteralPath $agentsPath) {
    $existing = [IO.File]::ReadAllText($agentsPath)
    if ($existing.Contains('<!-- rebacca-explicit-entry -->')) {
        Write-Warning 'An existing rebacca router was preserved. Check that its skill path is current.'
    } else {
        [IO.File]::AppendAllText($agentsPath, $router, $utf8)
    }
} else {
    [IO.File]::WriteAllText($agentsPath, $router, $utf8)
}
$manifest = Join-Path $destinationSkill 'SKILL.md'
if (-not (Test-Path -LiteralPath $manifest)) { throw 'Skill manifest verification failed.' }
$policy = [IO.File]::ReadAllText((Join-Path $destinationSkill 'agents\openai.yaml'))
if (-not $policy.Contains('allow_implicit_invocation: false')) { throw 'Explicit-only policy verification failed.' }
Write-Output "Installed: $manifest"
Write-Output "Configured tracker: $resolvedTracker"
Write-Output 'Configuration and file installation verified. Confirm discovery in Codex; restart Codex if it does not appear.'
