$ENABLE_ALL = 1             # "Enable all macros (not recommended; potentially dangerous code can run)"
$DISABLE_WITH_NOTIFY = 2    # "Disable all macros with notification"
$DISABLE_NOT_SIGNED = 3     # "Disable all macros except digitally signed macros"
$DISABLE_NO_NOTIFY = 4      # "Disable all macros without notification"

function Set-VBAWarningsFor($appPath, $val) {
    Write-Output "Writing VBAWarnings to " $appPath

    # Verify the path exists:
    if (!(Test-Path $appPath)) {
        Write-Output "Path does not exist; creating path " $appPath
        New-Item -Path $appPath -Force | Out-Null
        return
    }

    # Set the value
    Set-ItemProperty -Path $path -Name VBAWarnings -Value $val
}

$excel16 = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$word16 = "HKCU:\Software\Microsoft\Office\16.0\Word\Security"
$outlook16 = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security"

Set-VBAWarningsFor $excel16 $DISABLE_NO_NOTIFY
Set-VBAWarningsFor $word16 $DISABLE_NO_NOTIFY
Set-VBAWarningsFor $outlook16 $DISABLE_NO_NOTIFY