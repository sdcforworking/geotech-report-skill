# geotech-report 스킬 설치 (깃 없이)
# 사용: PowerShell에 아래 한 줄을 붙여 넣는다.
#   irm https://raw.githubusercontent.com/sdcforworking/geotech-report-skill/main/install.ps1 | iex
# 하는 일: GitHub에서 ZIP 받기 → %USERPROFILE%\.claude\skills\geotech-report 에 풀기
#          (이미 있으면 geotech-report.bak-날짜시각 으로 옮겨 두고, 지우지 않음) → 파이썬 패키지 설치
# 설치 위치를 바꾸려면 먼저 $env:GEOTECH_SKILL_DIR 에 폴더 경로를 넣는다.

& {
    $ErrorActionPreference = 'Stop'
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $zipUrl = 'https://github.com/sdcforworking/geotech-report-skill/archive/refs/heads/main.zip'
    $dest = if ($env:GEOTECH_SKILL_DIR) { $env:GEOTECH_SKILL_DIR } else { Join-Path $env:USERPROFILE '.claude\skills\geotech-report' }
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ('geotech-skill-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        Write-Host '1/3 스킬 내려받는 중...'
        $zip = Join-Path $tmp 'skill.zip'
        Invoke-WebRequest -Uri $zipUrl -OutFile $zip -UseBasicParsing
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [IO.Compression.ZipFile]::ExtractToDirectory($zip, (Join-Path $tmp 'x'))  # .NET은 zip의 UTF-8 이름 표시를 따름(한글 폴더명)
        $src = Get-ChildItem (Join-Path $tmp 'x') -Directory | Select-Object -First 1
        if (-not (Test-Path (Join-Path $src.FullName 'SKILL.md'))) { throw '받은 파일에 SKILL.md가 없습니다.' }

        Write-Host "2/3 설치: $dest"
        if (Test-Path $dest) {
            $bak = $dest + '.bak-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
            Move-Item $dest $bak
            Write-Host "    원래 있던 폴더는 옮겨 둠: $bak"
        }
        New-Item -ItemType Directory -Force -Path (Split-Path $dest) | Out-Null
        Move-Item $src.FullName $dest

        Write-Host '3/3 파이썬 패키지 설치'
        $py = $null
        foreach ($c in 'py', 'python') {
            $cmd = Get-Command $c -ErrorAction SilentlyContinue
            if ($cmd -and $cmd.Source -notlike '*WindowsApps*') { $py = $cmd.Source; break }  # WindowsApps = 스토어 안내용 가짜 python
        }
        if ($py) {
            $ErrorActionPreference = 'Continue'  # pip 경고(stderr)가 PowerShell 5.1에서 오류로 바뀌어 멈추지 않게
            & $py -m pip install openpyxl pymupdf pywin32 ezdxf pillow zstandard
            if ($LASTEXITCODE -ne 0) { Write-Warning 'pip 설치가 실패했습니다. 위 메시지를 확인하세요.' }
        } else {
            Write-Warning 'Python을 찾지 못했습니다. https://www.python.org 에서 3.10 이상을 깐 뒤 아래를 실행하세요:'
            Write-Host '    python -m pip install openpyxl pymupdf pywin32 ezdxf pillow zstandard'
        }
        Write-Host ''
        Write-Host '설치 끝. Claude Code를 다시 열면 geotech-report 스킬이 잡힙니다.'
        Write-Host '참고: user-CLAUDE.md 는 만든 사람의 개인 규칙입니다. 자기 CLAUDE.md 에 그대로 복사하지 마세요.'
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}
