#
# This will export some data from Active Directory into special phonebook file for Yealink phones.
# Phones then can download it automatically. (c) by Sasa.
#

$FolderLocation = "C:\inetpub\wwwroot-dfsr\voip\"
$strName = "company-phonebook"
$stream = [System.IO.StreamWriter] "$FolderLocation\\$strName.xml"
$stream.WriteLine("<YealinkIPPhoneDirectory>")

$root = [ADSI]"LDAP://dc=company,dc=local"
$filter = "(&(&(&(telephoneNumber=*)(!(msExchHideFromAddressLists=TRUE)))))"
$props = 'name','telephoneNumber','mobile'
$searcher = New-Object System.DirectoryServices.DirectorySearcher($root,$filter,$props)
$searcher.Sort.PropertyName = "telephoneNumber"
$searcher.Sort.Direction = [System.DirectoryServices.SortDirection]::Ascending
$searcher.FindAll() | ForEach-Object{ 
  $entry = $_.GetDirectoryEntry()
  $name = $entry.name
  $telephoneNumber = $entry.telephoneNumber
  $mobile = $entry.mobile
  $mobile = $mobile -replace "^."
  $stream.WriteLine("  <DirectoryEntry>")
  $stream.WriteLine("    <Name>"+ $name +"</Name>")
  $stream.WriteLine("    <Telephone>"+ $telephoneNumber +"</Telephone>")
  $stream.WriteLine("    <Telephone>"+ $mobile +"</Telephone>")
  $stream.WriteLine("  </DirectoryEntry>")
}

$stream.WriteLine("</YealinkIPPhoneDirectory>")
$stream.close()
