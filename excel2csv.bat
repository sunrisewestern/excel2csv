@if (@X)==(@Y) @end /* JScript comment
@echo off

set "inputFile=%~1"
set "tempFile=%~dpn1_temp.csv"
set "outputFile=%~dp1//SampleSheet.csv"
set "temp=%~dpn1_temp2.csv"

call :JSCRIPT


(for /f "tokens=1* delims=:" %%a in ('findstr /n .* "%tempFile%"') do (
    if %%a GTR 20 echo %%b
)
) > "%temp%"

(for /f "tokens=1* delims=:" %%a in ('findstr /n .* "%tempFile%"') do (
    if %%a LEQ 20 echo %%b
)
findstr /v ",,,,,,,,," "%temp%" 
) > "%outputFile%"

del "%temp%"
del "%tempFile%"

goto :EOF

:JSCRIPT
  @REM execute self as WSH JScript
  @cscript //nologo //E:jscript "%~f0" "%inputFile%" "%tempFile%"
  exit /b 
*/ {}
function wscriptMain(filename,tempname){
WScript.Echo(filename);
var objFSO = WScript.CreateObject("Scripting.FileSystemObject");
var path = objFSO.GetParentFolderName(filename);
WScript.Echo(path);
var xlCSVUTF8 = 62;

var objExcel = WScript.CreateObject("Excel.Application");
var objWorkbook = objExcel.Workbooks.Open(filename);
objExcel.DisplayAlerts = false;
objExcel.Visible = false;

var objWorksheet = objWorkbook.Worksheets("SampleSheet")
objWorksheet.SaveAs(  tempname, xlCSVUTF8);
objExcel.Quit();
}

wscriptMain(WScript.Arguments(0),WScript.Arguments(1));

