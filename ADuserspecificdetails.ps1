$importcsv = $null
$importcsv = import-csv C:\Temp\userlist.csv
$date = $(Get-Date -Format dd_MM_yyyy)
function get_userallinfo{

param($userupn)

begin{

$useralldata = $null
$useralldata = @()

}

process{

$useralldata = Get-ADUser -Filter {UserPrincipalName -eq $userupn} -Properties *


}
end{

return $useralldata

}
}

function export_csv{

param($exportobject, $export_filepath, $export_filename)

begin{

$path = $export_filepath + $export_filename

}

process{

$exportobject | Export-Csv -Path $path -Append -NoTypeInformation

}



}

for($u = 0; $u -lt $importcsv.upn.count ; $u++ )
{
$user, $useralldata, $whencreated, $office = $null
$user = $importcsv[$u].upn

$useralldata = get_userallinfo -userupn $user
if($useralldata -ne $null)
{
$whencreated = $useralldata.Created
$office = $useralldata.Office
if($office -eq $null)
{

$office = "Office Not Found"

}

if($whencreated -eq $null)
{

$whencreated = "Creation Date Not Found"

}
}
else{
if($useralldata -eq $null)
{
$whencreated = "Not Found"
$office = "Not Found"
}


}

$aduserobject = New-Object -TypeName psobject
$aduserobject | Add-Member -MemberType NoteProperty -Name upn -Value $user
$aduserobject | Add-Member -MemberType NoteProperty -Name created -Value $whencreated
$aduserobject | Add-Member -MemberType NoteProperty -Name office -Value $office

export_csv -exportobject $aduserobject -export_filepath "C:\Temp\" -export_filename "aduserdetails_$date.csv"

}