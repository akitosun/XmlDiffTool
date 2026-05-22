@echo off
setlocal

set "ROOT=%~dp0"
set "PROJECT=%ROOT%XmlDiffTool\XmlDiffTool.csproj"
set "OUTPUT=%ROOT%bin"

echo Building XmlDiffTool...
echo Output: %OUTPUT%

dotnet publish "%PROJECT%" -c Release -o "%OUTPUT%" --self-contained false
if errorlevel 1 (
    echo.
    echo Build failed.
    exit /b 1
)

echo.
echo Build completed successfully.
echo Executable: %OUTPUT%\XmlDiffTool.exe

endlocal
