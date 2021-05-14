@echo %~dp0 directory
@Reg Add "HKCR\*\shell\WordConverter" /VE /D "Convert to Word2010 (.docx)" /F >Nul
@Reg Add "HKCR\*\shell\WordConverter\command" /VE /D "\"%~dp0\WordConverter.exe\" \"%%1\"" /F >Nul
@Reg Add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WordConverter.exe" /VE /D "\"%~dp0\WordConverter.exe\"" /F >Nul
@Reg Add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WordConverter.exe" /V "Path" /D "\"%~dp0\"" /F >Nul
@echo Converter has been loaded
@pause