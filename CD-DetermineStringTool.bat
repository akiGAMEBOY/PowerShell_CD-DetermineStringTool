@ECHO OFF
REM #################################################################################
REM # 処理名　｜CD-DetermineStringTool（起動用バッチ）
REM # 機能　　｜PowerShell起動用のバッチ
REM #--------------------------------------------------------------------------------
REM # 　　　　｜-
REM #################################################################################
ECHO *---------------------------------------------------------
ECHO *
ECHO *  CD-DetermineStringTool
ECHO *
ECHO *---------------------------------------------------------
ECHO.
ECHO.
SET RETURNCODE=0
powershell -NoProfile -ExecutionPolicy Unrestricted -File .\Main.ps1
SET RETURNCODE=%ERRORLEVEL%

ECHO.
ECHO 処理が終了しました。
ECHO いずれかのキーを押すとウィンドウが閉じます。
PAUSE > NUL
EXIT %RETURNCODE%
