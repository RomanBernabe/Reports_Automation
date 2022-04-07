SET pwd=%cd%

SET zip=\Get-Zip.ps1
SET ziproute=%pwd%%zip%
Powershell.exe -ExecutionPolicy Bypass -File "%ziproute%"

::Note how we encased in quotes the ziproute to feed it to Powershell's File parameter.




SET rename=\rename_files.py
SET renameroute="%pwd%%rename%"
python %renameroute%



SET report=\make_report.py
SET reportroute="%pwd%%report%"
python %reportroute%

