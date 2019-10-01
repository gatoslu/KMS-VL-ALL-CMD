@echo off

set _Debug=0
set "SysPath=%Windir%\System32"
if exist "%Windir%\Sysnative\reg.exe" (set "SysPath=%Windir%\Sysnative")
set "Path=%SysPath%;%Windir%;%SysPath%\Wbem;%SysPath%\WindowsPowerShell\v1.0\"
fsutil dirty query %systemdrive% >nul 2>&1 || (
set "msg=ERROR: right click on the script and 'Run as administrator'"
goto :end
)

set xOS=x64
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" (if "%PROCESSOR_ARCHITEW6432%"=="" set xOS=x86)
set "_tempdir=%SystemRoot%\Temp"
set "_logpath=%~dpn0"
set "_workdir=%~dp0"
if "%_workdir:~-1%"=="\" set "_workdir=%_workdir:~0,-1%"

if %_Debug% EQU 0 (
  set "_Nul_1=1>nul"
  set "_Nul_2=2>nul"
  set "_Nul_2e=2^>nul"
  set "_Nul_1_2=1>nul 2>nul"
  call :Begin
) else (
  set "_Nul_1="
  set "_Nul_2="
  set "_Nul_2e="
  set "_Nul_1_2="
  echo.
  echo Running in Debug Mode...
  echo The window will be closed when finished
  call setlocal EnableDelayedExpansion
  copy /y nul "!_workdir!\#.rw" 1>nul 2>nul && (if exist "!_workdir!\#.rw" del /f /q "!_workdir!\#.rw") || (set "_logpath=!_tempdir!\%~n0")
  @echo on
  @prompt $G
  @call :Begin >"!_logpath!.tmp" 2>&1 &cmd /u /c type "!_logpath!.tmp">"!_logpath!_Debug.log"&del "!_logpath!.tmp"
)
exit /b

:Begin
color 1F
title Office 2013 Click2Run Retail2Volume
for /f "tokens=6 delims=[]. " %%# in ('ver') do set winbuild=%%#
set "_SLMGR=%SystemRoot%\System32\slmgr.vbs"

if %winbuild% lss 7601 (
set "msg=Windows 7 SP1 is the minimum supported OS..."
goto :end
)
sc query ClickToRunSvc %_Nul_1_2%
set error1=%errorlevel%
sc query OfficeSvc %_Nul_1_2%
set error2=%errorlevel%
if %error1% equ 1060 if %error2% equ 1060 (
set "msg=Could not detect Office ClickToRun service..."
goto :end
)

set _Office15=0
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul_2e%') do if exist "%%b\root\Licenses\*.xrm-ms" (
  set _Office15=1
)
if exist "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" (
  set _Office15=1&set "_OSPP=%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" (
  set _Office15=1&set "_OSPP=%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS"
)
if %_Office15% equ 0 (
set "msg=No installed Office 2013 product detected..."
goto :end
)

for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul_2e%') do if not errorlevel 1 (set "_InstallRoot=%%b\root")
if not "%_InstallRoot%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v PackageGUID" %_Nul_2e%') do if not errorlevel 1 (set "_GUID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds" %_Nul_2e%') do if not errorlevel 1 (set "ProductIds=%%b")
  set "_Config=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds"
  set "_PRIDs=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs"
  set "_OSPPReady=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"
) else (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul_2e%') do if not errorlevel 1 (set "_InstallRoot=%%b\root")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun /v PackageGUID" %_Nul_2e%') do if not errorlevel 1 (set "_GUID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds" %_Nul_2e%') do if not errorlevel 1 (set "ProductIds=%%b")
  set "_Config=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds"
  set "_PRIDs=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs"
  set "_OSPPReady=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration"
)
set "_OSPPReadT=REG_SZ"
if "%ProductIds%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid" %_Nul_2e%') do if not errorlevel 1 (set "ProductIds=%%b")
  set "_Config=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid"
  set "_OSPPReady=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun"
  set "_OSPPReadT=REG_DWORD"
)
set "_LicensesPath=%_InstallRoot%\Licenses"
set "_Integrator=%_InstallRoot%\integration\integrator.exe"

if "%ProductIds%"=="" (
set "msg=Could not detect Office ProductIDs..."
goto :end
)
setlocal EnableDelayedExpansion
if not exist "!_LicensesPath!\*.xrm-ms" (
set "msg=Could not detect Office Licenses files..."
goto :end
)
:: if not exist "!_Integrator!" (
:: set "msg=Could not detect Office Licenses Integrator..."
:: goto :end
:: )
if %winbuild% lss 9200 if not exist "!_OSPP!" (
set "msg=Could not detect Licensing tool OSPP.vbs..."
goto :end
)

:Check
echo.
echo ============================================================
echo Checking Office Licenses...
echo ============================================================
if %winbuild% geq 9200 (
set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService
set "_Integ=%_SLMGR% /ilc "
) else (
set spp=OfficeSoftwareProtectionProduct
set sps=OfficeSoftwareProtectionService
set _Integ="!_OSPP!" /inslic:
)
for /f "tokens=2 delims==" %%# in ('"wmic path %sps% get version /value" %_Nul_2e%') do set ver=%%#
wmic path %spp% where (Description like '%%KMSCLIENT%%' AND not LicenseFamily='Office15MondoR_KMS_Automation') get LicenseFamily %_Nul_2% | findstr /i /C:"Office" %_Nul_1% && (set _KMS=1) || (set _KMS=0)
wmic path %spp% where (Description like '%%TIMEBASED%%') get LicenseFamily %_Nul_2% | findstr /i /C:"Office" %_Nul_1% && (set _Time=1) || (set _Time=0)
wmic path %spp% where (Description like '%%Trial%%') get LicenseFamily %_Nul_2% | findstr /i /C:"Office" %_Nul_1% && (set _Time=1)
wmic path %spp% where (Description like '%%Grace%%') get LicenseFamily %_Nul_2% | findstr /i /C:"Office" %_Nul_1% && (set _Grace=1) || (set _Grace=0)
if %_Time% equ 0 if %_Grace% equ 0 if %_KMS% equ 1 (
set "msg=No Conversion or Cleanup Required..."
goto :GVLK
)

:Retail2Volume
echo.
echo ============================================================
echo Cleaning Current Office Licenses...
echo ============================================================
"!_workdir!\!xOS!\cleanospp.exe" -Licenses %_Nul_1_2%
echo.
echo ============================================================
echo Installing Office Volume Licenses...
echo ============================================================
echo.
for %%# in ("!_LicensesPath!\client-issuance-*.xrm-ms") do (
cscript //Nologo //B !_Integ!"%%#"
)
cscript //Nologo //B !_Integ!"!_LicensesPath!\pkeyconfig-office.xrm-ms"

set O15Ids=Standard,ProjectPro,VisioPro,ProjectStd,VisioStd,Access,Lync
set A15Ids=Excel,Groove,InfoPath,OneNote,Outlook,PowerPoint,Publisher,Word

echo %ProductIds%> "!_tempdir!\ProductIds.txt"
for %%a in (SPD,Mondo,O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem,Professional,HomeBusiness,HomeStudent,%O15Ids%,%A15Ids%,ProPlus) do (
set _%%a=0
)
for %%a in (SPD,Mondo,O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem,Professional,HomeBusiness,HomeStudent,%O15Ids%,%A15Ids%) do (
findstr /I /C:"%%aRetail" "!_tempdir!\ProductIds.txt" %_Nul_1% && set _%%a=1
)
wmic path %spp% get LicenseFamily > "!_tempdir!\sppchk.txt" 2>&1
for %%a in (Mondo,%O15Ids%,%A15Ids%) do (
findstr /I /C:"%%aVolume" "!_tempdir!\ProductIds.txt" %_Nul_1% && (
  find /i "%%aVL_KMS_Client" "!_tempdir!\sppchk.txt" %_Nul_1% && (set _%%a=0) || (set _%%a=1)
  )
)
reg query %_PRIDs%\Active\ProPlusRetail\x-none %_Nul_1_2% && set _ProPlus=1
reg query %_PRIDs%\Active\ProPlusRetail\x-none %_Nul_1_2% && (
find /i "OfficeProPlusVL_KMS_Client" "!_tempdir!\sppchk.txt" %_Nul_1% && (set _ProPlus=0) || (set _ProPlus=1)
)
del /f /q "!_tempdir!\sppchk.txt" >nul 2>&1
del /f /q "!_tempdir!\ProductIds.txt" >nul 2>&1

if !_Mondo! equ 1 (
echo Mondo Suite
echo.
call :InsLic Mondo
)
if !_O365ProPlus! equ 1 (
echo O365ProPlus Suite -^> Mondo Licenses
echo.
call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if !_Mondo! equ 0 call :InsLic Mondo
)
if !_O365SmallBusPrem! equ 1 if !_O365ProPlus! equ 0 (
set _O365ProPlus=1
echo O365SmallBusPrem Suite -^> Mondo Licenses
echo.
call :InsLic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! equ 0 call :InsLic Mondo
)
if !_O365HomePrem! equ 1 if !_O365SmallBusPrem! equ 0 if !_O365ProPlus! equ 0 (
set _O365ProPlus=1
echo O365HomePrem Suite -^> Mondo Licenses
echo.
call :InsLic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! equ 0 call :InsLic Mondo
)
if !_O365Business! equ 1 if !_O365HomePrem! equ 0 if !_O365SmallBusPrem! equ 0 if !_O365ProPlus! equ 0 (
set _O365ProPlus=1
echo O365Business Suite -^> Mondo Licenses
echo.
call :InsLic O365Business MCPBN-CPY7X-3PK9R-P6GTT-H8P8Y
if !_Mondo! equ 0 call :InsLic Mondo
)
if !_Mondo! equ 1 if !_O365ProPlus! equ 0 (
call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
goto :GVLK
)
if !_SPD! equ 1 if !_Mondo! equ 0 if !_O365ProPlus! equ 0 (
echo SharePointDesigner App -^> Mondo Licenses
echo.
call :InsLic Mondo
)
if !_ProPlus! equ 1 if !_O365ProPlus! equ 0 (
echo ProPlus Suite -^> ProPlus Licenses
echo.
call :InsLic ProPlus
)
if !_Professional! equ 1 if !_O365ProPlus! equ 0 if !_ProPlus! equ 0 (
echo Professional Suite -^> ProPlus Licenses
echo.
call :InsLic ProPlus
)
if !_Standard! equ 1 if !_O365ProPlus! equ 0 if !_ProPlus! equ 0 if !_Professional! equ 0 (
echo Standard Suite -^> Standard Licenses
echo.
call :InsLic Standard
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a! equ 1 (
echo %%a SKU -^> %%a Licenses
echo.
call :InsLic %%a
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus! equ 0 if !_Professional! equ 0 if !_Standard! equ 0 (
  set _Standard=1
  echo %%a Suite -^> Standard Licenses
  echo.
  call :InsLic Standard
  )
)
for %%a in (%A15Ids%) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus! equ 0 if !_Professional! equ 0 if !_Standard! equ 0 (
  echo %%a App
  echo.
  call :InsLic %%a
  )
)
for %%a in (Access) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus! equ 0 if !_Professional! equ 0 (
  echo %%a App
  echo.
  call :InsLic %%a
  )
)
for %%a in (Lync) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus! equ 0 (
  echo SkypeforBusiness App
  echo.
  call :InsLic %%a
  )
)
goto :GVLK

:InsLic
set "_ID=%1Volume"
set "_patt=%1VL_"
set "_pkey="
if not "%2"=="" (
set "_ID=%1Retail"
set "_patt=%1R_"
set "_pkey=%2"
)
reg delete %_OSPPReady% /f /v %_ID%.OSPPReady %_Nul_1_2%
for %%# in ("!_LicensesPath!\!_patt!*.xrm-ms") do (
cscript //Nologo //B !_Integ!"%%#"
)
if defined _pkey wmic path %sps% where version='%ver%' call InstallProductKey ProductKey="%_pkey%" %_Nul_1_2%
:: set "_pkey=PidKey=%2"
:: "!_Integrator!" /I /License PRIDName=%_ID% %_pkey% PackageGUID="%_GUID%" PackageRoot="!_InstallRoot!" %_Nul_1%
reg add %_OSPPReady% /f /v %_ID%.OSPPReady /t %_OSPPReadT% /d 1 %_Nul_1%
reg query %_Config% | findstr /I "%_ID%" %_Nul_1%
if %errorlevel% neq 0 (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Config%') do reg add %_Config% /t REG_SZ /d "%%b,%_ID%" /f %_Nul_1%
)
exit /b

:GVLK
echo ============================================================
echo Installing Missing KMS Client Keys...
echo ============================================================
echo.
for /f "tokens=2 delims==" %%# in ('"wmic path %spp% where (Description like 'Office 15, VOLUME_KMSCLIENT%%' AND PartialProductKey=NULL) get ID /value" %_Nul_2e%') do (set "app=%%#"&call :InsKey)
if exist "%SystemRoot%\System32\spp\store_test\2.0\tokens.dat" (
echo.
echo ============================================================
echo Refreshing Windows Insider Preview Licenses...
echo ============================================================
echo.
cscript //Nologo //B %_SLMGR% /rilc
)
set "msg=Finished"
goto :end

:InsKey
set "key="
for /f "tokens=2 delims==" %%# in ('"wmic path %spp% where ID='%app%' get LicenseFamily /value"') do echo %%#
call "!_workdir!\x86\keyOff.cmd" %app%
if "%key%" equ "" (echo Could not find matching gVLK&echo.&exit /b)
wmic path %sps% where version='%ver%' call InstallProductKey ProductKey="%key%" %_Nul_1_2%
set ERRORCODE=%ERRORLEVEL%
if %ERRORCODE% neq 0 (
cmd /c exit /b %ERRORCODE%
echo Failed: 0x!=ExitCode!
)
echo.
exit /b

:end
echo.
echo ============================================================
echo %msg%
echo ============================================================
echo.
echo Press any key to exit...
if %_Debug% EQU 0 pause >nul
goto :eof
