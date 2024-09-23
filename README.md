These are PowerShell scripts to convert .xls files to .xlsx and .doc files to .docx. After each .xls or .doc file is converted, the origical copy is moved into a new subfolder named "old" within the directory where the file was found. An output file will be written to the specified path. A password value of $null is passed for files that are password protected. This will prevent the application from being loaded and prompting for a password, but the protected files will not be converted.

PowerShell 7 also requires explicit references to the interoperability dlls for Word and Excel. If you do not have admin access, you will need to copy these dlls from c:\windows\assembly to a folder you have full access to.<br/><br/>
For the Excel script, you'll either need to copy Microsoft.Office.Interop.Excel.dll into the same path as the script or reference the path under c:\windows\assembly<br/>
For the Word script, you'll either need to copy Microsoft.Office.Interop.Word.dll into the same path as the script or reference the path under c:\windows\assembly

When run, the scripts will prompt for a filepath to traverse in search of .doc/.xls files.
