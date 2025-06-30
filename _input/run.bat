chcp 65001
@echo off

set "XML_CONFIG_FILE=${Env:ProgramFiles}\ITR\DocListHandler\input\DirToMonitor.xml"
set "DOC_LIST_TEMPLATE=${Env:ProgramFiles}\ITR\DocListHandler\input\DocListTemplate.xlsm"
set "LOG_DIR=$Env:USERPROFILE\Desktop\"

@echo on
"C:\Program Files\ITR\DocListManager\DocListManager.exe" --xmlConfigFile "%XML_CONFIG_FILE%" --docListTemplate "%DOC_LIST_TEMPLATE%" --logdir "%LOG_DIR%"
echo "Check the log file generated at" "%LOG_DIR%"

pause
