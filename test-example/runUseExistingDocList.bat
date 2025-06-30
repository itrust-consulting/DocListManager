@echo off

set "XML_CONFIG_FILE=.\_inputdoclist\DirToMonitorWithUseExistingDocList.xml"
set "DOC_LIST_TEMPLATE=.\_inputdoclist\DocListTemplate.xlsm"
set "LOG_DIR=."

@echo on
"C:\Program Files\ITR\DocListManager\DocListManager.exe" --xmlConfigFile "%XML_CONFIG_FILE%" --docListTemplate "%DOC_LIST_TEMPLATE%" --logdir "%LOG_DIR%"
echo "Check the log file generated at" "%LOG_DIR%"

pause
