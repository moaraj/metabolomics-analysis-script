setwd("D:/Dropbox/Aging BXD Study/D _ Metabolomics/D _ Protocol Optimization/CV Calcs")

excel_script <- 
    'if WScript.Arguments.Count < 2 Then
    WScript.Echo "Please specify the source and the destination files. Usage: ExcelToCsv <xls/xlsx source file> <csv destination file>"
Wscript.Quit
End If

csv_format = 6

Set objFSO = CreateObject("Scripting.FileSystemObject")

src_file = objFSO.GetAbsolutePathName(Wscript.Arguments.Item(0))
dest_file = objFSO.GetAbsolutePathName(WScript.Arguments.Item(1))

Dim oExcel
Set oExcel = CreateObject("Excel.Application")

Dim oBook
Set oBook = oExcel.Workbooks.Open(src_file)

oBook.SaveAs dest_file, csv_format

oBook.Close False
oExcel.Quit'

script_file_name = "ExcelToCsv.vbs"
write(excel_script,file = script_file_name)
# The script above allows command line conversion of xlxs file to csv conversiton in command line
#The script syntax: 
#XlsToCsv.vbs [sourcexlsFile].xls [destinationcsvfile].csv

library(tools)
abs_path <- file_path_as_absolute(dir(pattern = "\\.xls")[1])
cmd_command <- paste(c(script_file_name, abs_path, 
                       paste(strsplit(abs_path,".xls"),".csv", sep = "")), 
                     sep = " ", collapse = " ")

system(command = cmd_command)
