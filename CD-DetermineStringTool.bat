@ECHO OFF
REM #################################################################################
REM # �������@�bCD-DetermineStringTool�i�N���p�o�b�`�j
REM # �@�\�@�@�bPowerShell�N���p�̃o�b�`
REM #--------------------------------------------------------------------------------
REM # �@�@�@�@�b-
REM #################################################################################
ECHO *---------------------------------------------------------
ECHO *
ECHO *  CD-DetermineStringTool
ECHO *
ECHO *---------------------------------------------------------
ECHO.
ECHO.

powershell -NoProfile -ExecutionPolicy Unrestricted -File .\Main.ps1
SET RETURNCODE=%ERRORLEVEL%

ECHO.
ECHO �������I�����܂����B
ECHO �����ꂩ�̃L�[�������ƃE�B���h�E�����܂��B
PAUSE > NUL
EXIT %RETURNCODE%
