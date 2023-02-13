@if (@X)==(@Y) @end /* JScript comment
    @echo off

    cscript //E:JScript //nologo "%~f0" %*

    exit /b %errorlevel%

@if (@X)==(@Y) @end JScript comment */



var ARGS = WScript.Arguments;
WScript.Echo(ARGS.Item(0));
var objFSO = WScript.CreateObject("Scripting.FileSystemObject");
var path = objFSO.GetParentFolderName(ARGS(0));
WScript.Echo(path);
var xlCSVUTF8 = 62;

var objExcel = WScript.CreateObject("Excel.Application");
var objWorkbook = objExcel.Workbooks.Open(ARGS.Item(0));
objExcel.DisplayAlerts = false;
objExcel.Visible = false;

var objWorksheet = objWorkbook.Worksheets("SampleSheet")
objWorksheet.SaveAs(  path + "\\SampleSheet.csv", xlCSVUTF8);

objExcel.Quit();