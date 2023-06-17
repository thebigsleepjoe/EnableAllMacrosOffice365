$ENABLE_ALL = 1
# $ASK_FIRST = 2 # Included for completeness, but not used here

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

Set-VBAWarningsFor $excel16 $ENABLE_ALL
Set-VBAWarningsFor $word16 $ENABLE_ALL
Set-VBAWarningsFor $outlook16 $ENABLE_ALL