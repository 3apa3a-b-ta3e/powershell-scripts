#
# Checks if there is reboot pending due installation of updates, reboot PC if yes.
#

function Test-RebootRequired 
{
    $result = @{
        CBSRebootPending =$false
        WindowsUpdateRebootRequired = $false
    }

    #Check CBS Registry
    $key = Get-ChildItem "HKLM:Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction Ignore
    if ($key -ne $null) 
    {
        $result.CBSRebootPending = $true
    }
   
    #Check Windows Update
    $key = Get-Item "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction Ignore
    if($key -ne $null) 
    {
        $result.WindowsUpdateRebootRequired = $true
    }

    #Return Reboot required
    return $result.ContainsValue($true)
}


if(Test-RebootRequired)
{
    shutdown.exe /r /t 600 -c "WARNING! This PC will be automatically restarted in 10 minutes!"
}
