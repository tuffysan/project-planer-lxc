param(
    [string]$Version = "",
    [string]$Repo = "tuffysan/project-planer-lxc",
    [string]$Branch = "main",
    [switch]$NoRelease
)

$ErrorActionPreference = "Stop"

# PowerShell 7 can promote native stderr into ErrorRecords when
# PSNativeCommandUseErrorActionPreference is enabled. Git often writes
# normal progress/info to stderr, so disable that behavior for this script.
if (Get-Variable -Name PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue) {
    $PSNativeCommandUseErrorActionPreference = $false
}

function Fail {
    param([string]$Message)
    Write-Host ""
    Write-Host "ERROR: $Message" -ForegroundColor Red
    exit 1
}

function Invoke-Native {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [string[]]$ArgumentList = @(),
        [switch]$IgnoreExitCode,
        [switch]$Quiet
    )

    $display = if ($ArgumentList.Count -gt 0) {
        "$FilePath " + ($ArgumentList -join " ")
    } else {
        $FilePath
    }

    if (-not $Quiet) {
        Write-Host "> $display" -ForegroundColor DarkGray
    }

    # Temporarily relax PowerShell error handling so stderr from native tools
    # does not terminate the script. We trust the process exit code instead.
    $oldEap = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        & $FilePath @ArgumentList
        $code = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $oldEap
    }

    if (-not $IgnoreExitCode -and $code -ne 0) {
        Fail "Command failed with exit code $code`: $display"
    }

    return $code
}

function Invoke-NativeCapture {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [string[]]$ArgumentList = @()
    )

    $oldEap = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $output = & $FilePath @ArgumentList 2>&1
        $code = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $oldEap
    }

    return @{
        ExitCode = $code
        Output = @($output)
    }
}

function Find-Gh {
    $cmd = Get-Command gh -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $known = @(
        "$env:ProgramFiles\GitHub CLI\gh.exe",
        "${env:ProgramFiles(x86)}\GitHub CLI\gh.exe",
        "$env:LOCALAPPDATA\Programs\GitHub CLI\gh.exe"
    )
    foreach ($p in $known) {
        if ($p -and (Test-Path $p)) { return $p }
    }
    return $null
}

function Ensure-Gh {
    $gh = Find-Gh
    if ($gh) { return $gh }

    Write-Host ""
    Write-Host "GitHub CLI (gh) is not installed." -ForegroundColor Yellow

    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        Fail "GitHub CLI is missing and winget is not available. Install GitHub CLI from https://cli.github.com/ and run PUBLISH.cmd again."
    }

    Write-Host "Installing GitHub CLI automatically with winget..." -ForegroundColor Yellow
    Invoke-Native -FilePath $winget.Source -ArgumentList @(
        "install","--id","GitHub.cli","--exact",
        "--accept-package-agreements","--accept-source-agreements"
    )

    $env:Path += ";$env:ProgramFiles\GitHub CLI"
    $gh = Find-Gh
    if (-not $gh) {
        Fail "GitHub CLI was installed but gh.exe could not be found. Close this terminal, open a new PowerShell window and run PUBLISH.cmd again."
    }
    return $gh
}

function Ensure-GhAuth {
    param([Parameter(Mandatory=$true)][string]$GhPath)

    # GitHub CLI gives GH_TOKEN/GITHUB_TOKEN precedence over stored credentials.
    # A stale/invalid environment token can therefore block interactive login.
    $hadGhToken = -not [string]::IsNullOrWhiteSpace($env:GH_TOKEN)
    $hadGithubToken = -not [string]::IsNullOrWhiteSpace($env:GITHUB_TOKEN)

    $status = Invoke-NativeCapture -FilePath $GhPath -ArgumentList @("auth","status")
    if ($status.ExitCode -ne 0) {
        if ($hadGhToken -or $hadGithubToken) {
            Write-Host ""
            Write-Host "A GH_TOKEN/GITHUB_TOKEN environment variable is set but GitHub authentication failed." -ForegroundColor Yellow
            Write-Host "Clearing the token for this publish process so GitHub CLI can use browser login." -ForegroundColor Yellow

            # Clear only in this PowerShell process. User/machine environment is untouched.
            Remove-Item Env:GH_TOKEN -ErrorAction SilentlyContinue
            Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue

            # Re-check in case stored gh credentials are already valid once env tokens are gone.
            $status = Invoke-NativeCapture -FilePath $GhPath -ArgumentList @("auth","status")
        }

        if ($status.ExitCode -ne 0) {
            Write-Host ""
            Write-Host "GitHub authentication is required." -ForegroundColor Yellow
            Write-Host "A browser window will open for GitHub login." -ForegroundColor Cyan
            Invoke-Native -FilePath $GhPath -ArgumentList @(
                "auth","login",
                "--hostname","github.com",
                "--git-protocol","https",
                "--web"
            )
        }
    }

    # Final verification after login/token cleanup.
    $finalStatus = Invoke-NativeCapture -FilePath $GhPath -ArgumentList @("auth","status")
    if ($finalStatus.ExitCode -ne 0) {
        foreach ($line in $finalStatus.Output) {
            if ($line) { Write-Host $line -ForegroundColor DarkYellow }
        }
        Fail "GitHub authentication is still not valid."
    }

    Write-Host "Configuring Git to use GitHub CLI credentials..." -ForegroundColor Yellow
    Invoke-Native -FilePath $GhPath -ArgumentList @("auth","setup-git")
}

function Find-Python {
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        return @{ File = $py.Source; Args = @("-3") }
    }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        return @{ File = $python.Source; Args = @() }
    }

    return $null
}

function Get-NextPatchVersion {
    param([Parameter(Mandatory=$true)][string]$CurrentVersion)

    $parts = $CurrentVersion.Split(".")
    if ($parts.Count -ne 3) {
        Fail "Cannot auto-increment version '$CurrentVersion'. Expected MAJOR.MINOR.PATCH."
    }

    $major = [int]$parts[0]
    $minor = [int]$parts[1]
    $patch = [int]$parts[2] + 1
    return "$major.$minor.$patch"
}

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

# v2.1.2 no longer uses GitHub Actions. A workflow file from v2.1.2 would
# require a PAT with workflow scope, so remove any stale copy before staging.
$LegacyWorkflow = Join-Path $Root ".github\workflows\validate.yml"
if (Test-Path $LegacyWorkflow) {
    Remove-Item $LegacyWorkflow -Force
}

Write-Host "=== Project Planer LXC Publish ===" -ForegroundColor Cyan
Write-Host "Folder: $Root" -ForegroundColor DarkGray
Write-Host ""

$gitCmd = Get-Command git -ErrorAction SilentlyContinue
if (-not $gitCmd) {
    Fail "Git is not installed or not available in PATH."
}
$Git = $gitCmd.Source

if (-not $Version) {
    $VersionFile = Join-Path $Root "VERSION"
    if (-not (Test-Path $VersionFile)) {
        Fail "VERSION file is missing."
    }
    $Version = (Get-Content $VersionFile -Raw).Trim()
}

$Version = $Version.TrimStart("v")
$Tag = "v$Version"

if ($Version -notmatch '^\d+\.\d+\.\d+$') {
    Fail "Version must use semantic format, for example 2.1.2"
}

Set-Content -Path (Join-Path $Root "VERSION") -Value $Version -Encoding ascii

$Required = @(
    "install-lxc.sh",
    "install-app.sh",
    "update-lxc.sh",
    "requirements.txt",
    "README.md",
    "app\app.py",
    "scripts\project-plan.service"
)
foreach ($item in $Required) {
    if (-not (Test-Path (Join-Path $Root $item))) {
        Fail "Required file is missing: $item"
    }
}

Write-Host "[1/10] Validating project..." -ForegroundColor Yellow
$python = Find-Python
if ($python) {
    $pythonArgs = @()
    $pythonArgs += $python.Args
    $pythonArgs += @("-m", "py_compile", (Join-Path $Root "app\app.py"))
    $pyResult = Invoke-NativeCapture -FilePath $python.File -ArgumentList $pythonArgs
    if ($pyResult.ExitCode -ne 0) {
        Fail "Local Python validation failed. Fix the Python error before publishing."
        foreach ($line in $pyResult.Output) { Write-Host $line -ForegroundColor DarkYellow }
    }
} else {
    Write-Host "Python was not found. Skipping local Python syntax validation." -ForegroundColor Yellow
}

Write-Host "[2/10] Preparing GitHub authentication..." -ForegroundColor Yellow
$Gh = Ensure-Gh
Ensure-GhAuth -GhPath $Gh

Write-Host "[3/10] Checking Git repository..." -ForegroundColor Yellow
if (-not (Test-Path (Join-Path $Root ".git"))) {
    Invoke-Native -FilePath $Git -ArgumentList @("init")
}

$userName = (Invoke-NativeCapture -FilePath $Git -ArgumentList @("config","user.name"))
if ($userName.ExitCode -ne 0 -or -not ($userName.Output -join "").Trim()) {
    Invoke-Native -FilePath $Git -ArgumentList @("config","user.name","Andreas Nilsson")
}
$userEmail = (Invoke-NativeCapture -FilePath $Git -ArgumentList @("config","user.email"))
if ($userEmail.ExitCode -ne 0 -or -not ($userEmail.Output -join "").Trim()) {
    Invoke-Native -FilePath $Git -ArgumentList @("config","user.email","andreas.nilsson@users.noreply.github.com")
}

$currentRemote = Invoke-NativeCapture -FilePath $Git -ArgumentList @("remote","get-url","origin")
if ($currentRemote.ExitCode -ne 0 -or -not ($currentRemote.Output -join "").Trim()) {
    Invoke-Native -FilePath $Git -ArgumentList @("remote","add","origin","https://github.com/$Repo.git")
} elseif (($currentRemote.Output -join "`n") -notmatch [regex]::Escape($Repo)) {
    Write-Host "Updating origin to https://github.com/$Repo.git"
    Invoke-Native -FilePath $Git -ArgumentList @("remote","set-url","origin","https://github.com/$Repo.git")
}

Write-Host "[4/10] Testing GitHub repository access..." -ForegroundColor Yellow
Invoke-Native -FilePath $Gh -ArgumentList @("repo","view",$Repo) -Quiet

Write-Host "[5/10] Synchronizing with GitHub..." -ForegroundColor Yellow
$fetch = Invoke-NativeCapture -FilePath $Git -ArgumentList @("fetch","origin",$Branch)
$fetchOk = ($fetch.ExitCode -eq 0)

if ($fetch.Output.Count -gt 0) {
    foreach ($line in $fetch.Output) {
        if ($line) { Write-Host $line -ForegroundColor DarkGray }
    }
}

if ($fetchOk) {
    $headCheck = Invoke-NativeCapture -FilePath $Git -ArgumentList @("rev-parse","--verify","HEAD")
    $localHeadExists = ($headCheck.ExitCode -eq 0)

    if (-not $localHeadExists) {
        Invoke-Native -FilePath $Git -ArgumentList @("checkout","-B",$Branch,"origin/$Branch")
    } else {
        Invoke-Native -FilePath $Git -ArgumentList @("checkout","-B",$Branch)

        $merge = Invoke-NativeCapture -FilePath $Git -ArgumentList @(
            "merge","origin/$Branch",
            "--allow-unrelated-histories",
            "--no-edit"
        )
        foreach ($line in $merge.Output) {
            if ($line) { Write-Host $line -ForegroundColor DarkGray }
        }
        if ($merge.ExitCode -ne 0) {
            Write-Host ""
            Write-Host "Git could not merge the remote branch automatically." -ForegroundColor Red
            Write-Host "Run 'git status', resolve conflicts, then run PUBLISH.cmd again." -ForegroundColor Red
            exit 1
        }
    }
} else {
    Write-Host "Remote branch '$Branch' could not be fetched. Continuing with a new local branch." -ForegroundColor Yellow
    Invoke-Native -FilePath $Git -ArgumentList @("checkout","-B",$Branch)
}

Write-Host "[6/10] Building release package..." -ForegroundColor Yellow
$Dist = Join-Path $Root "dist"
if (Test-Path $Dist) { Remove-Item $Dist -Recurse -Force }
New-Item -ItemType Directory -Path $Dist | Out-Null

$ZipName = "project-planer-lxc-$Tag.zip"
$ZipPath = Join-Path $Dist $ZipName
$ExcludeTop = @(".git", "dist", "__pycache__")
$Items = Get-ChildItem $Root -Force | Where-Object { $ExcludeTop -notcontains $_.Name }
Compress-Archive -Path $Items.FullName -DestinationPath $ZipPath -CompressionLevel Optimal -Force

Write-Host "[7/10] Committing source to $Branch..." -ForegroundColor Yellow

# If v2.1.2 previously staged/committed dist locally, untrack it now.
# The release ZIP remains on disk and will still be uploaded by gh release.
$untrackDist = Invoke-NativeCapture -FilePath $Git -ArgumentList @("rm","-r","--cached","--ignore-unmatch","dist")
if ($untrackDist.ExitCode -ne 0) {
    Fail "Could not remove dist from Git tracking."
}

Invoke-Native -FilePath $Git -ArgumentList @("add","-A")

$status = Invoke-NativeCapture -FilePath $Git -ArgumentList @("status","--porcelain")
if (($status.Output -join "").Trim()) {
    Invoke-Native -FilePath $Git -ArgumentList @("commit","-m","Release $Tag")
} else {
    Write-Host "No source changes to commit."
}

Write-Host "[8/10] Pushing $Branch to GitHub..." -ForegroundColor Yellow
Invoke-Native -FilePath $Git -ArgumentList @("push","-u","origin",$Branch)

Write-Host "[9/10] Validating installer on GitHub..." -ForegroundColor Yellow
$RawUrl = "https://raw.githubusercontent.com/$Repo/$Branch/install-lxc.sh"
$Ok = $false
for ($i=1; $i -le 15; $i++) {
    try {
        $resp = Invoke-WebRequest -Uri $RawUrl -UseBasicParsing -TimeoutSec 20
        if ($resp.StatusCode -eq 200 -and $resp.Content -match '#!/usr/bin/env bash') {
            $Ok = $true
            break
        }
    } catch {}
    Write-Host "Waiting for GitHub raw content... ($i/15)" -ForegroundColor DarkGray
    Start-Sleep -Seconds 2
}
if (-not $Ok) {
    Fail "install-lxc.sh is still not reachable on GitHub main: $RawUrl"
}

if ($NoRelease) {
    Write-Host "[10/10] Release creation skipped (-NoRelease)." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Publish complete." -ForegroundColor Green
    Write-Host "Installer:"
    Write-Host "VERSION=$Version bash -c `"`$(curl -fsSL $RawUrl)`""
    exit 0
}

Write-Host "[10/10] Creating GitHub release..." -ForegroundColor Yellow

# If the requested tag/release already exists, automatically increment the patch
# version until a free version is found.
$TagAlreadyExistsWithoutRelease = $false

while ($true) {
    $releaseView = Invoke-NativeCapture -FilePath $Gh -ArgumentList @("release","view",$Tag,"--repo",$Repo)
    $remoteTag = Invoke-NativeCapture -FilePath $Gh -ArgumentList @(
        "api","repos/$Repo/git/ref/tags/$Tag"
    )
    $localTag = Invoke-NativeCapture -FilePath $Git -ArgumentList @(
        "show-ref","--tags","--verify","--quiet","refs/tags/$Tag"
    )

    if ($releaseView.ExitCode -ne 0 -and $remoteTag.ExitCode -eq 0) {
        Write-Host "Tag $Tag already exists but no GitHub Release exists. Reusing the tag and creating the missing release." -ForegroundColor Yellow
        $TagAlreadyExistsWithoutRelease = $true
        break
    }

    if ($releaseView.ExitCode -ne 0 -and $remoteTag.ExitCode -ne 0 -and $localTag.ExitCode -ne 0) {
        break
    }

    $oldVersion = $Version
    $Version = Get-NextPatchVersion -CurrentVersion $Version
    $Tag = "v$Version"
    Set-Content -Path (Join-Path $Root "VERSION") -Value $Version -Encoding ascii
    Write-Host "Version v$oldVersion already exists. Using $Tag instead." -ForegroundColor Yellow
}

# If auto-increment changed VERSION after the source commit, commit the updated
# VERSION file before tagging/releasing.
$versionStatus = Invoke-NativeCapture -FilePath $Git -ArgumentList @("status","--porcelain","VERSION")
if (($versionStatus.Output -join "").Trim()) {
    Invoke-Native -FilePath $Git -ArgumentList @("add","VERSION")
    Invoke-Native -FilePath $Git -ArgumentList @("commit","-m","Set release version $Tag")
    Invoke-Native -FilePath $Git -ArgumentList @("push","origin",$Branch)
}

if (-not $TagAlreadyExistsWithoutRelease) {

    Invoke-Native -FilePath $Git -ArgumentList @("tag","-a",$Tag,"-m","Release $Tag")

    Invoke-Native -FilePath $Git -ArgumentList @("push","origin",$Tag)

}

# Rebuild release ZIP in case the patch version was auto-incremented.
$FinalZipName = "project-planer-lxc-$Tag.zip"
$FinalZipPath = Join-Path $Dist $FinalZipName
if (Test-Path $FinalZipPath) {
    Remove-Item $FinalZipPath -Force
}
$Items = Get-ChildItem $Root -Force | Where-Object { $ExcludeTop -notcontains $_.Name }
Compress-Archive -Path $Items.FullName -DestinationPath $FinalZipPath -CompressionLevel Optimal -Force
$ZipPath = $FinalZipPath

$Notes = @"

Project Planer LXC $Tag



Install on Proxmox:



VERSION=$Version bash -c "`$(curl -fsSL https://raw.githubusercontent.com/$Repo/$Branch/install-lxc.sh)"

"@



$NotesFile = Join-Path $env:TEMP "project-planer-release-notes-$Version.md"

Set-Content -Path $NotesFile -Value $Notes -Encoding utf8

try {

    Invoke-Native -FilePath $Gh -ArgumentList @(

        "release","create",$Tag,$ZipPath,

        "--repo",$Repo,

        "--title",$Tag,

        "--notes-file",$NotesFile,

        "--verify-tag"

    )

}

finally {

    Remove-Item $NotesFile -Force -ErrorAction SilentlyContinue

}

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "PUBLISH COMPLETE" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host "Release: https://github.com/$Repo/releases/tag/$Tag"
Write-Host ""
Write-Host "Install on Proxmox:"
Write-Host "VERSION=$Version bash -c `"`$(curl -fsSL $RawUrl)`""
Write-Host ""
