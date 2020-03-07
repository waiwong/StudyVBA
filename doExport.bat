SET curDir=%~dp0
ECHO %curDIr%
:: ExportVBAfromXLS.vbs %curDIr%studyVBA.xlsm
ExportVBAfromXLS.vbs %curDIr%ConsolidateCSV.xlsm
:: CSCRIPT //d ExportVBAfromXLS.vbs %curDIr%ConsolidateCSV.xlsm