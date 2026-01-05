function Get-ScheduledTaskComHandler {
    $Path = "$($ENV:windir)\System32\Tasks"
    
    # Ensure the registry drive is mapped
    if (!(Get-PSDrive HKCR -ErrorAction SilentlyContinue)) {
        New-PSDrive -PSProvider Registry -Root HKEY_CLASSES_ROOT -Name HKCR | Out-Null
    }

    $Files = Get-ChildItem -Path $Path -Recurse | Where-Object { ! $_.PSIsContainer }
    
    $Results = foreach ($File in $Files) {
        try {
            $XML = [xml](Get-Content $File.FullName -ErrorAction Stop)
            if ($XML.Task.Actions.ComHandler) {
                $CLSID = $XML.Task.Actions.ComHandler.ClassID
                
                # Try to find the DLL associated with this CLSID
                $RegPath = "HKCR:\CLSID\$CLSID\InprocServer32"
                $DllPath = (Get-ItemProperty -LiteralPath $RegPath -ErrorAction SilentlyContinue).'(default)'

                [PSCustomObject]@{
                    TaskName = $File.Name
                    CLSID    = $CLSID
                    DllPath  = $DllPath
                }
            }
        } catch {
            # Skip files we can't read
        }
    }
    return $Results
}

# RUN THE FUNCTION AND SAVE TO FILE
$MyResults = Get-ScheduledTaskComHandler
$MyResults | Out-GridView                                     # Opens a searchable window
$MyResults | Export-Csv -Path "ComHandlers.csv" -NoTypeInformation

Write-Host "Found $($MyResults.Count) handlers. Saved to Documents\ComHandlers.csv" -ForegroundColor Green
