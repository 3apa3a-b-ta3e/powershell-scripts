#
# Here we will add an "AssignPrimaryToken" privilege to the local Administrators group.
# Sometimes it can be required to successfully deploy from Jenkins jobs.
# Also this privilege can be added to any other group determined by SID.
# This is different from assigning via GPO because here we didn`t delete any already added users/groups (which is impossible to do with GPO).
#

$tmp = [System.IO.Path]::GetTempFileName()
secedit /export /cfg "$tmp.inf" | Out-Null
(gc -Encoding ascii "$tmp.inf") -replace '^SeAssignPrimaryTokenPrivilege .+', "`$0,*S-1-5-32-544" | sc -Encoding ascii "$tmp.inf"
secedit /import /cfg "$tmp.inf" /db "$tmp.sdb" | Out-Null
secedit /configure /db "$tmp.sdb" | Out-Null
rm $tmp* -ea 0
