function Get-ScheduledTaskComHandler {
    $Path = "$($ENV:windir)\System32\Tasks"
    
    if (!(Get-PSDrive HKCR -ErrorAction SilentlyContinue)) {
        New-PSDrive -PSProvider Registry -Root HKEY_CLASSES_ROOT -Name HKCR | Out-Null
    }

    $Files = Get-ChildItem -Path $Path -Recurse | Where-Object { ! $_.PSIsContainer }
    
    $Results = foreach ($File in $Files) {
        try {
            $XML = [xml](Get-Content $File.FullName -ErrorAction Stop)
            
            # Check if this task uses a ComHandler action
            if ($XML.Task.Actions.ComHandler) {
                $CLSID = $XML.Task.Actions.ComHandler.ClassID
                
                # 1. GET DLL PATH
                $RegPath = "HKCR:\CLSID\$CLSID\InprocServer32"
                $DllPath = (Get-ItemProperty -LiteralPath $RegPath -ErrorAction SilentlyContinue).'(default)'

                # 2. GET DESCRIPTION (Using If/Else)
                if ($XML.Task.RegistrationInfo.Description) {
                    $Description = $XML.Task.RegistrationInfo.Description
                } else {
                    $Description = "No description provided"
                }

                # 3. GET RUNTYPE / TRIGGERS (Using If/Else)
                if ($XML.Task.Triggers) {
                    # This joins multiple triggers into a single string (e.g., "LogonTrigger, TimeTrigger")
                    $Triggers = ($XML.Task.Triggers.ChildNodes | ForEach-Object { $_.Name }) -join ", "
                } else {
                    $Triggers = "Manual / No Trigger"
                }

                [PSCustomObject]@{
                    TaskName    = $File.Name
                    CLSID       = $CLSID
                    DllPath     = $DllPath
                    RunType     = $Triggers
                    Description = $Description
                }
            }
        } catch {
            # Catch XML parsing errors or access denied
        }
    }
    return $Results
}

# Execution
$MyResults = Get-ScheduledTaskComHandler
$MyResults | Out-GridView
$MyResults | Export-Csv -Path "$HOME\Documents\ComHandlers.csv" -NoTypeInformation

Write-Host "Found $($MyResults.Count) handlers. Saved to Documents\ComHandlers.csv" -ForegroundColor Green
