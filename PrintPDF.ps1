$adobe='C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe'
$printername='DocuCentre-V C2263 (3d:e2:01)'
$File = $args[0]
$arglist='/S /T "{0}" "{1}"' -f "$File", $printername

start-Process $adobe -ArgumentList $arglist

$sleepcount = $CountFiles.Count + [int]11
Start-Sleep $sleepcount

