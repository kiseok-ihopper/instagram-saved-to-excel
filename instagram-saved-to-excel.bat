@echo off
chcp 65001 1>NUL
setlocal enabledelayedexpansion

echo =============================================
echo  인스타그램 저장목록 HTML → 엑셀 변환기
echo =============================================
echo.

if not "%~1"=="" (
    set "HTML_FILE=%~1"
    echo [입력] %~1
) else (
    set "HTML_FILE="
    for %%f in ("%~dp0*.html") do (
        if "!HTML_FILE!"=="" set "HTML_FILE=%%~ff"
    )
    if "!HTML_FILE!"=="" (
        echo [오류] HTML 파일을 찾을 수 없습니다.
        echo.
        echo  사용 방법:
        echo   1. HTML 파일을 이 bat 파일 위로 드래그앤드롭
        echo   2. 또는 이 bat 파일과 같은 폴더에 HTML 파일 위치
        echo.
        pause
        exit /b 1
    )
    echo [자동감지] !HTML_FILE!
)

python --version 2>&1 1>NUL
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo  https://www.python.org/downloads/ 에서 설치 후 재실행하세요.
    echo.
    pause
    exit /b 1
)

echo.
echo [패키지 확인 중...]
python -c "import subprocess,sys; r=subprocess.run([sys.executable,'-c','import openpyxl'],capture_output=True); sys.exit(r.returncode)"
if errorlevel 1 (
    echo  openpyxl 설치 중...
    pip install openpyxl -q
)

if not "%~1"=="" (
    set "OUT_FILE=%~dpn1_변환결과.xlsx"
) else (
    for %%f in ("!HTML_FILE!") do set "OUT_FILE=%%~dpnf_변환결과.xlsx"
)

set "INSTA_HTML=!HTML_FILE!"
set "INSTA_OUT=!OUT_FILE!"

echo.
echo [변환 중...]
echo.

python -c "import base64,os; exec(base64.b64decode('IyAtKi0gY29kaW5nOiB1dGYtOCAtKi0KaW1wb3J0IHJlLCBzeXMsIG9zCmltcG9ydCBvcGVucHl4bApmcm9tIG9wZW5weXhsLnN0eWxlcyBpbXBvcnQgRm9udCwgUGF0dGVybkZpbGwsIEFsaWdubWVudAoKSFRNTF9QQVRIID0gb3MuZW52aXJvblsnSU5TVEFfSFRNTCddCk9VVF9QQVRIICA9IG9zLmVudmlyb25bJ0lOU1RBX09VVCddCgpwcmludCgnSFRNTCDtjIzsnbwg7J2964qUIOykkS4uLicpCndpdGggb3BlbihIVE1MX1BBVEgsICdyJywgZW5jb2Rpbmc9J3V0Zi04JykgYXMgZjoKICAgIGNvbnRlbnQgPSBmLnJlYWQoKQoKcHJpbnQoJ+2MjOyLsSDspJEuLi4nKQpoMl9wb3NpdGlvbnMgPSBbbS5zdGFydCgpIGZvciBtIGluIHJlLmZpbmRpdGVyKHInPGgyW14+XSo+JywgY29udGVudCldCmgyX3Bvc2l0aW9ucy5hcHBlbmQobGVuKGNvbnRlbnQpKQoKY2F0X3BhdHRlcm4gID0gcmUuY29tcGlsZShyJ05hbWU8ZGl2PjxkaXY+KFtePF0rKTwvZGl2PjwvZGl2Pi4qP0NyZWF0aW9uIFRpbWU8L3RkPjx0ZFtePl0qPihbXjxdKyk8L3RkPi4qP1VwZGF0ZSBUaW1lPC90ZD48dGRbXj5dKj4oW148XSspPC90ZD4nLCByZS5ET1RBTEwpCnBvc3RfcGF0dGVybiA9IHJlLmNvbXBpbGUocidOYW1lPGRpdj48YSB0YXJnZXQ9Il9ibGFuayIgaHJlZj0iKFteIl0rKSI+KFtePF0rKTwvYT48L2Rpdj4uKj9BZGRlZCBUaW1lPC90ZD48dGRbXj5dKj4oW148XSspPC90ZD4nLCByZS5ET1RBTEwpCgpyb3dzID0gW10KZm9yIGkgaW4gcmFuZ2UobGVuKGgyX3Bvc2l0aW9ucykgLSAxKToKICAgIHNlY3Rpb24gPSBjb250ZW50W2gyX3Bvc2l0aW9uc1tpXTpoMl9wb3NpdGlvbnNbaSsxXV0KICAgIGNhdF9tYXRjaCA9IGNhdF9wYXR0ZXJuLnNlYXJjaChzZWN0aW9uKQogICAgaWYgbm90IGNhdF9tYXRjaDoKICAgICAgICBjb250aW51ZQogICAgY2F0ZWdvcnkgICAgPSBjYXRfbWF0Y2guZ3JvdXAoMSkuc3RyaXAoKQogICAgdXBkYXRlX3RpbWUgPSBjYXRfbWF0Y2guZ3JvdXAoMykuc3RyaXAoKQogICAgZm9yIHBvc3QgaW4gcG9zdF9wYXR0ZXJuLmZpbmRpdGVyKHNlY3Rpb24pOgogICAgICAgIHVybCAgICAgICAgPSBwb3N0Lmdyb3VwKDEpLnN0cmlwKCkKICAgICAgICBhY2NvdW50ICAgID0gcG9zdC5ncm91cCgyKS5zdHJpcCgpCiAgICAgICAgYWRkZWRfdGltZSA9IHBvc3QuZ3JvdXAoMykuc3RyaXAoKQogICAgICAgIHBvc3RfdHlwZSAgPSAi66a07IqkIiBpZiAiL3JlZWwvIiBpbiB1cmwgZWxzZSAi6rKM7Iuc6riAIiBpZiAiL3AvIiBpbiB1cmwgZWxzZSAi6riw7YOAIgogICAgICAgIHBvc3RfaWQgICAgPSBbcCBmb3IgcCBpbiB1cmwuc3BsaXQoIi8iKSBpZiBwXVstMV0KICAgICAgICByb3dzLmFwcGVuZCgoY2F0ZWdvcnksIGFjY291bnQsIHVwZGF0ZV90aW1lLCBhZGRlZF90aW1lLCB1cmwsIHBvc3RfdHlwZSwgcG9zdF9pZCkpCgpwcmludChmIuyImCB7bGVuKHJvd3MpfeqwnCDqsozsi5zquIAg7YyM7IuxIOyZhOujjCIpCgp3YiA9IG9wZW5weXhsLldvcmtib29rKCkKd3MgPSB3Yi5hY3RpdmUKd3MudGl0bGUgPSAi7KCA7J6l66qp66GdIgoKaGVhZGVycyAgICA9IFsi7Lm07YWM6rOg66asIiwgIuqzhOygleuqhSIsICLsl4XroZzrk5wg64Kg7KecIiwgIuyggOyepe2VnCDrgqDsp5wiLCAiVVJMIiwgIuq1rOu2hCIsICLqsozsi5zrrLwgSUQiXQpoZHJfZmlsbCAgID0gUGF0dGVybkZpbGwoInNvbGlkIiwgZmdDb2xvcj0iNDQ3MkM0IikKaGRyX2ZvbnQgICA9IEZvbnQoYm9sZD1UcnVlLCBjb2xvcj0iRkZGRkZGIiwgc2l6ZT0xMSkKY29sX3dpZHRocyA9IFsyMCwgMjUsIDIyLCAyMiwgNTUsIDEwLCAxOF0KCmZvciBjb2wsIGggaW4gZW51bWVyYXRlKGhlYWRlcnMsIDEpOgogICAgY2VsbCA9IHdzLmNlbGwocm93PTEsIGNvbHVtbj1jb2wsIHZhbHVlPWgpCiAgICBjZWxsLmZpbGwgPSBoZHJfZmlsbAogICAgY2VsbC5mb250ID0gaGRyX2ZvbnQKICAgIGNlbGwuYWxpZ25tZW50ID0gQWxpZ25tZW50KGhvcml6b250YWw9ImNlbnRlciIsIHZlcnRpY2FsPSJjZW50ZXIiKQp3cy5yb3dfZGltZW5zaW9uc1sxXS5oZWlnaHQgPSAyMApmb3IgY29sLCB3IGluIGVudW1lcmF0ZShjb2xfd2lkdGhzLCAxKToKICAgIHdzLmNvbHVtbl9kaW1lbnNpb25zW3dzLmNlbGwoMSwgY29sKS5jb2x1bW5fbGV0dGVyXS53aWR0aCA9IHcKCmZvciByb3dfaWR4LCAoY2F0ZWdvcnksIGFjY291bnQsIHVwZGF0ZV90aW1lLCBhZGRlZF90aW1lLCB1cmwsIHBvc3RfdHlwZSwgcG9zdF9pZCkgaW4gZW51bWVyYXRlKHJvd3MsIDIpOgogICAgd3MuY2VsbChyb3c9cm93X2lkeCwgY29sdW1uPTEsIHZhbHVlPWNhdGVnb3J5KQogICAgd3MuY2VsbChyb3c9cm93X2lkeCwgY29sdW1uPTIsIHZhbHVlPWFjY291bnQpCiAgICB3cy5jZWxsKHJvdz1yb3dfaWR4LCBjb2x1bW49MywgdmFsdWU9dXBkYXRlX3RpbWUpCiAgICB3cy5jZWxsKHJvdz1yb3dfaWR4LCBjb2x1bW49NCwgdmFsdWU9YWRkZWRfdGltZSkKICAgIGMgPSB3cy5jZWxsKHJvdz1yb3dfaWR4LCBjb2x1bW49NSwgdmFsdWU9dXJsKQogICAgYy5oeXBlcmxpbmsgPSB1cmwKICAgIGMuZm9udCA9IEZvbnQoY29sb3I9IjA1NjNDMSIsIHVuZGVybGluZT0ic2luZ2xlIikKICAgIHdzLmNlbGwocm93PXJvd19pZHgsIGNvbHVtbj02LCB2YWx1ZT1wb3N0X3R5cGUpCiAgICB3cy5jZWxsKHJvdz1yb3dfaWR4LCBjb2x1bW49NywgdmFsdWU9cG9zdF9pZCkKCndzLmZyZWV6ZV9wYW5lcyA9ICJBMiIKd3MuYXV0b19maWx0ZXIucmVmID0gd3MuZGltZW5zaW9ucwoKcHJpbnQoIuyggOyepSDspJEuLi4iKQp3Yi5zYXZlKE9VVF9QQVRIKQpwcmludChmIuyZhOujjDoge09VVF9QQVRIfSIpCg==').decode('utf-8'))"

set RESULT=%errorlevel%
echo.
if %RESULT%==0 (
    echo [성공] 저장 위치: !OUT_FILE!
    echo.
    start "" "!OUT_FILE!"
) else (
    echo [실패] 변환 중 오류가 발생했습니다.
)

echo.
pause
