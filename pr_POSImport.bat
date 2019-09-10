@echo off
setlocal enabledelayedexpansion

:: Read Config Values from Config File
for /f "delims=^= tokens=1,2 skip=1" %%G in (%~dp0\..\BatchInput\POS_Tool_Config.txt) do set %%G=%%H

:: Declare variables
SET MQL_EXE=%ENOVIAINSTALLPATH%\3DSpace\win_b64\code\bin\mql.exe
SET POSIMPORTBATCHPATH=%ENOVIAINSTALLPATH%\3DSpace\win_b64\code\command\VPLMPosImport.bat
SET PNO_EXCEL_TEMPLATE=%POSTOOLINSTALLPATH%\BatchInput\PnOStructure.xlsm

::Setting timestamp value for Log File name.
SET LOG_FOLDER_PATH=%POSTOOLINSTALLPATH%\Logs
SET TODAYS_LOGS=%LOG_FOLDER_PATH%\%DATE%
SET IMPORTINPUTFOLDERPATH=%POSTOOLINSTALLPATH%\ImportInput
SET TODAYS_IMPORTINPUTFOLDER=%IMPORTINPUTFOLDERPATH%\%DATE%
SET IMPORTINPUTFILENAME=%TODAYS_IMPORTINPUTFOLDER%\PnOImportInput.txt
SET LOG_FILE_NAME=%TODAYS_LOGS%\POSImportToolLog.txt

:: create Today's folder in Logs and ImportInput folders
IF NOT EXIST %TODAYS_LOGS% (
	md "%TODAYS_LOGS%"
)

IF NOT EXIST %TODAYS_IMPORTINPUTFOLDER% (
	md "%TODAYS_IMPORTINPUTFOLDER%"
)

::  Get and format timestamp to add to Import Input and Log files after execution.
set TIMESTAMP=%DATE%%TIME%
SET TIMESTAMP=%TIMESTAMP:-=%
SET TIMESTAMP=%TIMESTAMP::=%
SET TIMESTAMP=%TIMESTAMP:.=%

SET ARGS[0]=%PNO_EXCEL_TEMPLATE%

:: MQL command to pass to MQL.EXE
SET MQL_COMMAND="set context user %MQL_ADMINUSER% pass %MQL_ADMINUSER_PSWD%; exec prog POSImport -method generateImportInput %PNO_EXCEL_TEMPLATE%;"

:: Execute POS tool through MQL command
call %MQL_EXE% -c %MQL_COMMAND%

:: Declare variables for Passport User Registration
SET REPORT_ARG=%TODAYS_LOGS%\POSImportOOTBLog.txt
SET PASSPORTINPUTFILE=%TODAYS_IMPORTINPUTFOLDER%\PassportInput.txt
SET PASSPORT_BATCH=%ENOVIAINSTALLPATH%\3DPassport\win_b64\code\command\PassportUserImport.bat

:: Take backup of Import Input and Log files.
if EXIST %IMPORTINPUTFILENAME% (
	call %POSIMPORTBATCHPATH% -server %POS_SERVERURL_ARG% -user %POS_USER_ARG% -password %POS_PASSWORD_ARG% -context %POS_CONTEXT_ARG% -file %IMPORTINPUTFILENAME% -report %REPORT_ARG%
	move /y %IMPORTINPUTFILENAME% %IMPORTINPUTFILENAME%_%TIMESTAMP%
)

if EXIST %PASSPORTINPUTFILE% (
	call %PASSPORT_BATCH% -url %PASSPORT_URL% -file %PASSPORTINPUTFILE%
	move /y %PASSPORTINPUTFILE% %PASSPORTINPUTFILE%_%TIMESTAMP%
)

if EXIST %LOG_FILE_NAME% (
	move /y %LOG_FILE_NAME% %LOG_FILE_NAME%_%TIMESTAMP%
)

if EXIST %REPORT_ARG% (
	move /y %REPORT_ARG% %REPORT_ARG%_%TIMESTAMP%
)
echo Tool Excecution is Complete...