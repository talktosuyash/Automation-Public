##########################################################################################  
# Powershell Script to Install Patches on list of server according to windows os version #
#                                                                                        #
# 1. Get Hotix details from server.                                                      #
# 2. check destination folder arch if not create it.                                     #
# 3. copy hotfix to destination server.                                                  #
# 4. install hotfix by extracting cab file.                                              #
##########################################################################################


$patchingservers = Import-Csv $csvpath
$sourcekbshare = ":\Updates"
$getkb = Get-ChildItem $sourcekbshare

function gethotfix{

param($kbcheck, $compname)

begin{

$hotfixoutput, $hotfixcheckoutput, $trimkb = $null
$hotfixoutput, $hotfixcheckoutput,$trimkb = @()
}

process{
$trimkb = $($kbcheck.split('-') | where {$_ -like 'kb*'}).trim()
$scriptblock = {


 Get-HotFix 

}
 
$hotfixoutput = Invoke-Command -ComputerName $compname -ScriptBlock $scriptblock

if($hotfixoutput.hotfixid.contains($trimkb.toupper()))
{

$hotfixcheckoutput = "Present"

}
else{

$hotfixcheckoutput = "Not Present"
}

}
end{

return $hotfixoutput, $hotfixcheckoutput, $trimkb
}

}

function foldercheck{

param($compname)

begin{



}
process{
$foldercreation = $null
$foldercreation = @()


if($(Get-ChildItem "\\$compname\C$\temp" -Directory).Name -contains "Patching")
{
Write-host "Folder Pesent" -BackgroundColor Black -ForegroundColor Yellow


}
else{

if($(Get-ChildItem "\\$compname\C$\temp" -Directory).Name -notcontains "Patching"){

Write-host "Folder Not Pesent" -BackgroundColor Black -ForegroundColor Red
$foldercreation = New-Item -ItemType Directory -Path "\\$compname\C$\temp" -Name Patching -Force -Verbose

}
}



}

end{

return $foldercreation
}
}

function installkb{

param($sourcepath, $destinationpath, $folder, $compname, $kbdata)

begin{
$process, $invokeop = $null
$process, $invokeop= @()

}

Process{

C:\Windows\System32\expand.exe -f:* $sourcepath $destinationpath

$getcabfile = Get-ChildItem $destinationpath | where {$_.Extension -eq ".cab"}

for($i = 0; $i -lt $getcabfile.count; $i++)
{

if($getcabfile[$i].Name -match $kbdata )
{
$cabfile = $getcabfile[$i].Name



$packagepath = "C:\temp\Patching\$folder\$cabfile"
$logfilepath = "C:\temp\Patching\$folder\$kbdata.log"
Write-Host "Installing $kbdata in $compname server" 
$kbinstalloutput = Invoke-Command -ComputerName $compname -ScriptBlock {Add-WindowsPackage -Online -PackagePath $using:packagepath -NoRestart -LogLevel WarningsInfo -LogPath $using:logfilepath}

if($kbinstalloutput.online -eq $true)
{
$hotfixmatch = gethotfixafterinstall -compname $compname -kbdata $kbdata
}
if($hotfixmatch.installedby.length -ne 0) 
{
$installedby = $hotfixmatch.installedby

}

if($hotfixmatch.installedby.length -eq 0) 
{
$installedby = "reboot pending"

}
if($hotfixmatch.installedon.length -ne 0) 
{
$installedon = $hotfixmatch.installedon

}

if($hotfixmatch.installedon.length -eq 0) 
{
$installedon = "reboot pending"

}

if($kbinstalloutput.RestartNeeded -eq $true)
{
$restartstatus = 'Restart Needed'

}
$hotfixmatchobject = $null
$hotfixmatchobject = New-Object -TypeName psobject
$hotfixmatchobject | Add-Member -MemberType NoteProperty -Name ServerName -Value $compname
$hotfixmatchobject | Add-Member -MemberType NoteProperty -Name Installationstatus -Value $hotfixmatch.hotfixid
$hotfixmatchobject | Add-Member -MemberType NoteProperty -Name description -Value $hotfixmatch.description
$hotfixmatchobject | Add-Member -MemberType NoteProperty -Name installedby -Value $installedby
$hotfixmatchobject | Add-Member -MemberType NoteProperty -Name installedon -Value $installedon
$hotfixmatchobject | Add-Member -MemberType NoteProperty -Name rebootstatus -Value $restartstatus
$hotfixmatchobject | Export-Csv .\PatchingOutput.csv -Append -NoTypeInformation
}


}
}


end{
return $kbinstalloutput, $hotfixmatch

}

}

function gethotfixafterinstall{

param($compname, $kbdata)

begin{
$hotfixmatch = $null
$hotfixmatch = @()

}

process{

$hotfixmatch = Invoke-Command -ComputerName $compname -ScriptBlock {Get-HotFix} | where {$_.hotfixid -eq $kbdata}

}

end{

return $hotfixmatch 
}
}

function createfolder{

param($kbname, $compname)

process{


New-Item -ItemType Directory -Path "\\$compname\C$\temp\Patching" -Name $kbname -Force -Verbose

}



}

function copykb{

param($compname, $patchlist)

begin{
 $computername = $compname
 $process,$invokeop =  @()
}


process{
$composinfo = Get-WMIObject -ComputerName $computername win32_operatingsystem -Credential $cred
$ostype = $composinfo.name.Split('|')[0]
$patchlist = $getkb

for($j = 0; $j -lt $patchlist.count; $j++)
{
$patch = $patchlist[$j]
$hotfixoutput, $hotfixcheckoutput,$trimkb = gethotfix -kbcheck $patch.Name -compname $compname
if($hotfixcheckoutput -eq "Not Present")
{
$source = $sourcekbshare + $($patch.Name)
$destination = "\\$computername\C$\Temp\Patching\$trimkb\$($patch.Name)"
$destination2 = "\\$computername\C$\Temp\Patching\$trimkb\"
$dest3 = "C:\Temp\Patching\$trimkb\$($patch.Name)"

if($ostype -match 2016)
{

$kbinstalloutput = $null
foldercheck -compname $compname
createfolder -kbname $trimkb -compname $compname
Copy-Item -Path $source -Destination $destination -Verbose
$kbinstalloutput =  installkb -sourcepath $destination -destinationpath $destination2 -folder $trimkb -compname $compname -kbdata $trimkb 

}

if($ostype -match 2019)
{
$kbinstalloutput = $null
foldercheck -compname $compname
createfolder -kbname $trimkb -compname $compname
Copy-Item -Path $source -Destination $destination -Verbose
$kbinstalloutput =  installkb -sourcepath $destination -destinationpath $destination2 -folder $trimkb -compname $compname -kbdata $trimkb

}

if($ostype -match 2012)
{
$kbinstalloutput = $null
foldercheck -compname $compname
createfolder -kbname $trimkb -compname $compname
Copy-Item -Path $source -Destination $destination -Verbose
$kbinstalloutput =  installkb -sourcepath $destination -destinationpath $destination2 -folder $trimkb -compname $compname -kbdata $trimkb

}

$objectcopy = New-Object -TypeName psobject
$objectcopy | Add-Member -MemberType NoteProperty -Name Machinename -Value $compname
$objectcopy | Add-Member -MemberType NoteProperty -Name Patch -Value $patch.Name
$objectcopy
}
}
}

end{

return $process,$invokeop 

}


}

for($i = 0 ; $i -lt $patchingservers.Name.count; $i++)
{

$server = $patchingservers[$i].Name
Write-Host "Installing KB on server $server..." -BackgroundColor Black -ForegroundColor Green
copykb -compname $server -patchlist $getkb

}