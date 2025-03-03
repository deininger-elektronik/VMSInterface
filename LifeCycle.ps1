
# Set up a LifeCycle managment for the Program as well!!

$programPath = "C:\ProgramData\VMSInterface\VMSInterface.exe"
$programName = "VMSInterface"

while ($true) {
    $process = Get-Process -Name $programName -ErrorAction SilentlyContinue
    if (!$process) {
        Start-Process -FilePath $programPath -ArgumentList "/silent" -Verb RunAs
    }
    Start-Sleep -Seconds 5
}