Start-Process -FilePath $args[0] -verb print

$Logfile = "./"+(Get-date -Format 'yyyy-MM-dd')+".txt"   #LogFile
Add-content $Logfile "$(Get-date -Format 'yyyy-MM-dd tthh:mm:ss') Print: $args"

Start-Sleep -s 5