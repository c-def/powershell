These are PowerShell scripts to convert .xls files to .xlsx and .doc files to .docx. After each .xls or .doc file is converted, the origical copy is moved into a new subfolder named "old" within the directory where the file was found. An output file will be written to the specified path.

PowerShell 7 is required to use the -Parallel flag with ForEach-Object.

PowerShell 7 also requires references to the interoperability dlls for Word and Excel. If you do not have admin access, you will need to copy these dlls from c:\windows\assembly to a folder you have full access to.
For the Excel script, you'll either need to copy Microsoft.Office.Interop.Excel.dll into the same path as the script or reference the path under c:\windows\assembly<br/>
For the Word script, you'll either need to copy Microsoft.Office.Interop.Word.dll into the same path as the script or reference the path under c:\windows\assembly

Set $folderpath at the top of the script. This is the path the script will recursively run through to convert documents.

I noticed an issue with the script moving files to the "old" directory if they had brackets in the filename. I've added two if statements to replace the brackets with parentheses.
