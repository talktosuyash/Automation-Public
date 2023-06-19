
cls
$count = $null

$refreshcsv, $imprefreshtoken, $impzohodetails, $refreshtoken,$refreshcsv, $rclientid, $rsecretid, $refreshheaders, $zohourl, $timediffernce = $null
$zohourl = "https://accounts.zoho.in/oauth/v2/token"
Write-Host "Importing Refresh token"

$mainpath = ".\manageengin"
$checkfolderpresent_important = Get-ChildItem $mainpath
if(($checkfolderpresent_important | where { $_.Name -contains "Important" }).Name -ne "Important" ){
try{
New-Item -Path '.\manageengin\zohoinfo' -ItemType Directory -ErrorAction SilentlyContinue
Write-Host "Zohoinfo Folder Created into '.\manageengin\'path please add zohodetails.csv file with following headers `n clientId,clientSecret,authCode-AllRequests,authCode-AllAssets,authCode-ALLCMDB with respective values" -ForegroundColor Yellow -BackgroundColor Black
break
}
catch{

Write-Host "zohoinfo Folder Already Present"
}
}

if(($checkfolderpresent_important | where { $_.Name -contains "Token CSVs" }).Name -ne "Token CSVs"){
try{
New-Item -Path '.\manageengin\Token CSVs' -ItemType Directory -ErrorAction SilentlyContinue
New-Item -Path '.\manageengin\Token CSVs\Assets' -ItemType Directory -ErrorAction SilentlyContinue
New-Item -Path '.\manageengin\Token CSVs\Requests' -ItemType Directory -ErrorAction SilentlyContinue
New-Item -Path '.\manageengin\Token CSVs\CMDB' -ItemType Directory -ErrorAction SilentlyContinue
}
catch{

Write-Host "Token CSVs Folder Already Present"
}
}


else{

Write-Host "Folder Structure present"
$assetrefreshtokenpath, $cmdbrefreshtokenpath, $requestrefreshtokenpath = $null

$zohopath = ".\manageengin\Important\zohoinfo\zohodetails.csv"
$assetrefreshtokenpath = ".\manageengin\Token CSVs\donotdelete"
$cmdbrefreshtokenpath = ".\manageengin\Token CSVs\donotdelete"
$requestrefreshtokenpath = ".\manageengin\Token CSVs\donotdelete"


$assetrefreshtokenpathinfo = ".\manageengin\Token CSVs\Assets"
$cmdbrefreshtokenpathinfo = ".\manageengin\Token CSVs\CMDB"
$requestrefreshtokenpathinfo = ".\manageengin\Token CSVs\Requests"

$assetrefreshcsv = Get-ChildItem $requestrefreshtokenpath | where {$_.Name -like "*_asset*"} | Sort-Object -Descending $_.CreationTime | Select-Object -First 1
$cmdbrefreshcsv = Get-ChildItem $requestrefreshtokenpath | where {$_.Name -like "*_cmdb*"} | Sort-Object -Descending $_.CreationTime | Select-Object -First 1
$requestrefreshcsv = Get-ChildItem $requestrefreshtokenpath | where {$_.Name -like "*_request*"} | Sort-Object -Descending $_.CreationTime | Select-Object -First 1

$imprefreshtoken_asset = import-csv "$($assetrefreshcsv.PSPath)"
$imprefreshtoken_request =import-csv "$($requestrefreshcsv.PSPath)"
$imprefreshtoken_cmdb = import-csv "$($cmdbrefreshcsv.PSPath)"
$impzohodetails = Import-Csv $zohopath

$refreshtoken_asset = $imprefreshtoken_asset.refreshtoken
$refreshtoken_request = $imprefreshtoken_request.refreshtoken
$refreshtoken_cmdb = $imprefreshtoken_cmdb.refreshtoken

$rclientid = $impzohodetails.clientId
$rsecretid = $impzohodetails.clientSecret

#################### ASSET ########################
$refeshbody, $refreshheaders, $newpresentdate, $timediffernce, $refeshbody, $tempolddate_asset, $tempdate = $null
$refeshbody = @{
    grant_type = "refresh_token"
    client_id = $rclientid
    client_secret = $rsecretid
    refresh_token = $refreshtoken_asset
}

$refreshheaders = @{
    "Content-Type" = "application/x-www-form-urlencoded"
}
$Global:tempolddate_asset #= @()
$tempdate = (Get-Date -Format 'dd/MM/yyyy hh:mm tt') 
$Global:tempolddate_asset = [datetime]::ParseExact($tempdate, "dd/MM/yyyy hh:mm tt", [System.Globalization.CultureInfo]::InvariantCulture);

$newassetresponce = $null
$newassetresponce = Invoke-RestMethod -Method POST -Uri $zohourl -Headers $refreshheaders -Body $refeshbody -ErrorAction SilentlyContinue


$newassetaccesstoken, $newassetrefreshtoken, $newaccesstokenobject = $null
Write-Host "timedifference is $($timediffernce.Minutes) mins equal to 3mins" -BackgroundColor Black -ForegroundColor green
$newassetaccesstoken = $newassetresponce.access_token
$newassetrefreshtoken = $refreshtoken_asset
$assetitems  = $null
$assetitems = Get-ChildItem '.\manageengin\Token CSVs\Assets'
if ($assetitems -ne $null)
{
Remove-Item $assetitems.PSPath -Verbose

$newaccesstokenobject = New-Object -TypeName psobject
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name accesstoken -Value $newassetaccesstoken
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name refreshtoken -Value $newassetrefreshtoken
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name time -Value $Global:tempolddate_asset 

Write-Host "Asset New Access Token : $newassetaccesstoken" -ForegroundColor Yellow -BackgroundColor Black
$newaccesstokenobject | Export-Csv "$assetrefreshtokenpathinfo\$($newassetaccesstoken).csv" -NoTypeInformation
}

################################### REQUEST ###########################

$refeshbody, $refreshheaders, $newpresentdate, $timediffernce, $refeshbody, $tempolddate_req, $tempdate = $null
$refeshbody = @{
    grant_type = "refresh_token"
    client_id = $rclientid
    client_secret = $rsecretid
    refresh_token = $refreshtoken_request
}

$refreshheaders = @{
    "Content-Type" = "application/x-www-form-urlencoded"
}

$Global:tempolddate_req #= @()
$Global:tempolddate_req #= $null
$tempdate = (Get-Date -Format 'dd/MM/yyyy hh:mm tt') 
$Global:tempolddate_req = [datetime]::ParseExact($tempdate, "dd/MM/yyyy hh:mm tt", [System.Globalization.CultureInfo]::InvariantCulture);

$newreqresponce = $null
$newreqresponce = Invoke-RestMethod -Method POST -Uri $zohourl -Headers $refreshheaders -Body $refeshbody -ErrorAction SilentlyContinue

Write-Host "timedifference is $($timediffernce.Minutes) mins equal to 3mins" -BackgroundColor Black -ForegroundColor green
$newreqaccesstoken = $newreqresponce.access_token
$newreqrefreshtoken = $refreshtoken_request
$requestitems = $null
$requestitems = Get-ChildItem '.\manageengin\Token CSVs\Requests'

if($requestitems -ne $null)
{
Remove-Item $requestitems.PSPath -Verbose

$newaccesstokenobject = New-Object -TypeName psobject
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name accesstoken -Value $newreqaccesstoken
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name refreshtoken -Value $newreqrefreshtoken
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name time -Value $Global:tempolddate_req

Write-Host "Request New Access Token : $newreqaccesstoken" -ForegroundColor Yellow -BackgroundColor Black
$newaccesstokenobject | Export-Csv "$requestrefreshtokenpathinfo\$($newreqaccesstoken).csv" -NoTypeInformation
}
###################################### CMDB #############################################
Write-Host "Inside CMDB checking" -ForegroundColor White -BackgroundColor Black
$refeshbody, $refreshheaders, $newpresentdate, $timediffernce, $refeshbody, $tempolddate_cmdb, $tempdate = $null
$refeshbody = @{
    grant_type = "refresh_token"
    client_id = $rclientid
    client_secret = $rsecretid
    refresh_token = $refreshtoken_cmdb
}

$refreshheaders = @{
    "Content-Type" = "application/x-www-form-urlencoded"
}

$tempdate = (Get-Date -Format 'dd/MM/yyyy hh:mm tt') 
$Global:tempolddate_cmdb #= @()
$Global:tempolddate_cmdb #= $null
$Global:tempolddate_cmdb = [datetime]::ParseExact($tempdate, "dd/MM/yyyy hh:mm tt", [System.Globalization.CultureInfo]::InvariantCulture);

$newcmdbresponce = $null
$newcmdbresponce = Invoke-RestMethod -Method POST -Uri $zohourl -Headers $refreshheaders -Body $refeshbody -ErrorAction SilentlyContinue

$newaccesstoken_cmdb, $newrefreshtoken, $newrefreshtoken_cmdb = $null
Write-Host "timedifference is $($timediffernce.Minutes) mins equal to 3mins" -BackgroundColor Black -ForegroundColor green
$newaccesstoken_cmdb = $newcmdbresponce.access_token
$newrefreshtoken_cmdb = $refreshtoken_cmdb

$cmdbitems = $null
$cmdbitems = Get-ChildItem '.\manageengin\Token CSVs\CMDB'
if($cmdbitems -ne $null)
{
Remove-Item $cmdbitems.PSPath -Verbose


$newaccesstokenobject = New-Object -TypeName psobject
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name accesstoken -Value $newaccesstoken_cmdb
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name refreshtoken -Value $newrefreshtoken_cmdb
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name time -Value $Global:tempolddate_cmdb

Write-Host "CMDB New Access Token : $newaccesstoken_cmdb" -ForegroundColor Yellow -BackgroundColor Black
$newaccesstokenobject | Export-Csv "$cmdbrefreshtokenpathinfo\$($newaccesstoken_cmdb).csv" -NoTypeInformation

}
}
$countertest=0

do {
$count=$count+1
$count
$refreshcsv, $imprefreshtoken, $impzohodetails, $refreshtoken,$refreshcsv, $rclientid, $rsecretid, $refreshheaders, $zohourl, $timediffernce = $null
$zohourl = "https://accounts.zoho.in/oauth/v2/token"
Write-Host "Importing Refresh token"
$zohopath = ".\manageengin\Important\zohoinfo\zohodetails.csv"
$assetrefreshtokenpath = ".\manageengin\Token CSVs\Assets"
$cmdbrefreshtokenpath = ".\manageengin\Token CSVs\CMDB"
$requestrefreshtokenpath = ".\manageengin\Token CSVs\Requests"

$assetrefreshcsv = Get-ChildItem $assetrefreshtokenpath | Sort-Object -Descending $_.CreationTime | Select-Object -First 1
$cmdbrefreshcsv = Get-ChildItem $cmdbrefreshtokenpath | Sort-Object -Descending $_.CreationTime | Select-Object -First 1
$requestrefreshcsv = Get-ChildItem $requestrefreshtokenpath | Sort-Object -Descending $_.CreationTime | Select-Object -First 1

$imprefreshtoken_asset = import-csv "$($assetrefreshcsv.PSPath)"
$imprefreshtoken_request =import-csv "$($requestrefreshcsv.PSPath)"
$imprefreshtoken_cmdb = import-csv "$($cmdbrefreshcsv.PSPath)"
$impzohodetails = Import-Csv $zohopath

Write-Host "Date is: $(Get-Date)" 
### Refresh Token
$refreshtoken_asset = $imprefreshtoken_asset.refreshtoken
$refreshtoken_request = $imprefreshtoken_request.refreshtoken
$refreshtoken_cmdb = $imprefreshtoken_cmdb.refreshtoken

$rclientid = $impzohodetails.clientId
$rsecretid = $impzohodetails.clientSecret

if($refreshtoken_asset -ne $null){
Write-Host "Inside Asset checking" -ForegroundColor White -BackgroundColor Black
$refeshbody, $refreshheaders, $newpresentdate, $timediffernce, $refeshbody = $null
$refeshbody = @{
    grant_type = "refresh_token"
    client_id = $rclientid
    client_secret = $rsecretid
    refresh_token = $refreshtoken_asset
}

$refreshheaders = @{
    "Content-Type" = "application/x-www-form-urlencoded"
}
$predate = Get-Date -Format 'dd/MM/yyyy hh:mm tt'
$newpresentdate = [datetime]::ParseExact($predate, "dd/MM/yyyy hh:mm tt", [System.Globalization.CultureInfo]::InvariantCulture);
$newpresentdate
Start-Sleep -Seconds 1
Write-Host $Global:tempolddate_asset -BackgroundColor Red
$timediffernce = New-TimeSpan -Start $newpresentdate -End $Global:tempolddate_asset

Write-Host $timediffernce -BackgroundColor Green
Start-Sleep -Seconds 1
if($timediffernce.Minutes -eq '-1'){
try{
$assetrefreshtokenpath, $assetrefreshcsv, $imprefreshtoken_asset, $refreshtoken_asset = $null
$assetrefreshtokenpath = ".\manageengin\Token CSVs\donotdelete"
$assetrefreshcsv = Get-ChildItem $assetrefreshtokenpath | where {$_.Name -like "*_asset*"} | Sort-Object -Descending $_.CreationTime | Select-Object -First 1
$imprefreshtoken_asset = import-csv "$($assetrefreshcsv.PSPath)"
$refreshtoken_asset = $imprefreshtoken_asset.refreshtoken 


Write-Host "Inside Asset checking" -ForegroundColor White -BackgroundColor Black
$refeshbody, $refreshheaders, $newpresentdate, $timediffernce, $refeshbody = $null
$refeshbodyasset = @{
    grant_type = "refresh_token"
    client_id = $rclientid
    client_secret = $rsecretid
    refresh_token = $refreshtoken_asset
}
$countertest=$countertest+1
Write-Host $countertest -BackgroundColor Green
Start-Sleep -Seconds 1

$newassetresponce = $null
$newassetresponce = Invoke-RestMethod -Method POST -Uri $zohourl -Headers $refreshheaders -Body $refeshbodyasset -ErrorAction SilentlyContinue
}
catch{

Write-Host "newassetresponce var is empty" -ForegroundColor Red -BackgroundColor White

}

if($newassetresponce -ne $null){
$newassetaccesstoken, $newassetrefreshtoken, $newaccesstokenobject = $null
Write-Host "timedifference is $($timediffernce.Minutes) mins equal to 3mins" -BackgroundColor Black -ForegroundColor green
$newassetaccesstoken = $newassetresponce.access_token
$newassetrefreshtoken = $refreshtoken_asset
#$Global:tempolddate_asset = @()
#$Global:tempolddate_asset = $null
$predate = Get-Date -Format 'dd/MM/yyyy hh:mm tt'
$newpresentdate = [datetime]::ParseExact($predate, "dd/MM/yyyy hh:mm tt", [System.Globalization.CultureInfo]::InvariantCulture);
$newpresentdate
$Global:tempolddate_asset =$newpresentdate
 
Write-Host $newpresentdate -BackgroundColor Green
Start-Sleep -Seconds 1

$assetitems  = $null
$assetitems = Get-ChildItem '.\manageengin\Token CSVs\Assets'


if ($assetitems -ne $null)
{
Remove-Item $assetitems.PSPath -Verbose 

$newaccesstokenobject = New-Object -TypeName psobject
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name accesstoken -Value $newassetaccesstoken
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name refreshtoken -Value $newassetrefreshtoken
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name time -Value $newpresentdate

Write-Host "Asset New Access Token : $newassetaccesstoken" -ForegroundColor Yellow -BackgroundColor Black
$newaccesstokenobject | Export-Csv "$assetrefreshtokenpathinfo\$($newassetaccesstoken).csv" -NoTypeInformation
}
}
}
}

if($refreshtoken_request -ne $null){
Write-Host "Inside Request checking" -ForegroundColor White -BackgroundColor Black
$refeshbody, $refreshheaders, $timediffernce, $refeshbody = $null
$refeshbody = @{
    grant_type = "refresh_token"
    client_id = $rclientid
    client_secret = $rsecretid
    refresh_token = $refreshtoken_request
}

$refreshheaders = @{
    "Content-Type" = "application/x-www-form-urlencoded"
}
$predate = Get-Date -Format 'dd/MM/yyyy hh:mm tt'
$newpresentdate = [datetime]::ParseExact($predate, "dd/MM/yyyy hh:mm tt", [System.Globalization.CultureInfo]::InvariantCulture);
$newpresentdate.ToString('dd-MM-yyyy hh:mm:ss tt')
$timediffernce = New-TimeSpan -Start $newpresentdate -End $Global:tempolddate_req

if($timediffernce.Minutes -eq '-1'){
try{
$requestrefreshtokenpath, $requestrefreshcsv, $imprefreshtoken_request = $null
$requestrefreshtokenpath = ".\manageengin\Token CSVs\donotdelete"
$requestrefreshcsv = Get-ChildItem $requestrefreshtokenpath | where {$_.Name -like "*_request*"} | Sort-Object -Descending $_.CreationTime | Select-Object -First 1
$imprefreshtoken_request =import-csv "$($requestrefreshcsv.PSPath)"
$refreshtoken_request = $imprefreshtoken_request.refreshtoken

$refeshbody, $refreshheaders, $timediffernce, $refeshbody = $null
$refeshbody_request = @{
    grant_type = "refresh_token"
    client_id = $rclientid
    client_secret = $rsecretid
    refresh_token = $refreshtoken_request
}


$newreqresponce = $null
$newreqresponce = Invoke-RestMethod -Method POST -Uri $zohourl -Headers $refreshheaders -Body $refeshbody_request -ErrorAction SilentlyContinue
}
catch{

Write-Host "Responce var is empty" -ForegroundColor Red -BackgroundColor White

}

if($newreqresponce -ne $null){
$newaccesstoken, $newrefreshtoken, $newaccesstokenobject = $null
Write-Host "timedifference is $($timediffernce.Minutes) mins equal to 3mins" -BackgroundColor Black -ForegroundColor green
$newreqaccesstoken = $newreqresponce.access_token
$newreqrefreshtoken = $refreshtoken_request
$Global:tempolddate_req #= @()
$Global:tempolddate_req #= $null
$predate = Get-Date -Format 'dd/MM/yyyy hh:mm tt'
$newpresentdate = [datetime]::ParseExact($predate, "dd/MM/yyyy hh:mm tt", [System.Globalization.CultureInfo]::InvariantCulture);
$newpresentdate
$Global:tempolddate_req = $newpresentdate
$requestitems = $null
$requestitems = Get-ChildItem '.\manageengin\Token CSVs\Requests'

if($requestitems -ne $null)
{
Remove-Item $requestitems.PSPath -Verbose 

$newaccesstokenobject = New-Object -TypeName psobject
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name accesstoken -Value $newreqaccesstoken
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name refreshtoken -Value $newreqrefreshtoken
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name time -Value $newpresentdate

Write-Host "Request New Access Token : $newreqaccesstoken" -ForegroundColor Yellow -BackgroundColor Black
$newaccesstokenobject | Export-Csv "$requestrefreshtokenpathinfo\$($newreqaccesstoken).csv" -NoTypeInformation
}
}
}
}

if($refreshtoken_cmdb -ne $null){
Write-Host "Inside CMDB checking" -ForegroundColor White -BackgroundColor Black
$refeshbody, $refreshheaders, $timediffernce, $refeshbody = $null
$refeshbody = @{
    grant_type = "refresh_token"
    client_id = $rclientid
    client_secret = $rsecretid
    refresh_token = $refreshtoken_cmdb
}

$refreshheaders = @{
    "Content-Type" = "application/x-www-form-urlencoded"
}
$predate = Get-Date -Format 'dd/MM/yyyy hh:mm tt'
$newpresentdate = [datetime]::ParseExact($predate, "dd/MM/yyyy hh:mm tt", [System.Globalization.CultureInfo]::InvariantCulture);
$newpresentdate
$timediffernce = New-TimeSpan -Start $newpresentdate -End $Global:tempolddate_cmdb

if($timediffernce.Minutes -eq '-1'){
try{
$cmdbrefreshtokenpath, $cmdbrefreshcsv, $imprefreshtoken_cmdb, $refreshtoken_cmdb = $null
$cmdbrefreshtokenpath = ".\manageengin\Token CSVs\donotdelete"
$cmdbrefreshcsv = Get-ChildItem $cmdbrefreshtokenpath | where {$_.Name -like "*_cmdb*"} | Sort-Object -Descending $_.CreationTime | Select-Object -First 1
$imprefreshtoken_cmdb = import-csv "$($cmdbrefreshcsv.PSPath)"
$refreshtoken_cmdb = $imprefreshtoken_cmdb.refreshtoken

Write-Host "Inside CMDB checking" -ForegroundColor White -BackgroundColor Black
$refeshbody_cmdb, $refreshheaders, $timediffernce, $refeshbody = $null
$refeshbody_cmdb = @{
    grant_type = "refresh_token"
    client_id = $rclientid
    client_secret = $rsecretid
    refresh_token = $refreshtoken_cmdb
}

$newcmdbresponce = $null
$newcmdbresponce = Invoke-RestMethod -Method POST -Uri $zohourl -Headers $refreshheaders -Body $refeshbody_cmdb -ErrorAction SilentlyContinue



}
catch{

Write-Host "Responce var is empty" -ForegroundColor Red -BackgroundColor White

}

if($newcmdbresponce -ne $null){
$newaccesstoken_cmdb, $newrefreshtoken, $newrefreshtoken_cmdb = $null
Write-Host "timedifference is $($timediffernce.Minutes) mins equal to 3mins" -BackgroundColor Black -ForegroundColor green
$newaccesstoken_cmdb = $newcmdbresponce.access_token
$newrefreshtoken_cmdb = $refreshtoken_cmdb

$Global:tempolddate_cmdb #= @()
$Global:tempolddate_cmdb #= $null
$predate = Get-Date -Format 'dd/MM/yyyy hh:mm tt'
$newpresentdate = [datetime]::ParseExact($predate, "dd/MM/yyyy hh:mm tt", [System.Globalization.CultureInfo]::InvariantCulture);
$newpresentdate
$Global:tempolddate_cmdb = $newpresentdate

$cmdbitems = $null
$cmdbitems = Get-ChildItem '.\manageengin\Token CSVs\CMDB'
if($cmdbitems -ne $null)
{
Remove-Item $cmdbitems.PSPath -Verbose

$newaccesstokenobject = New-Object -TypeName psobject
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name accesstoken -Value $newaccesstoken_cmdb
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name refreshtoken -Value $newrefreshtoken_cmdb
$newaccesstokenobject | Add-Member -MemberType NoteProperty -Name time -Value $newpresentdate

Write-Host "CMDB New Access Token : $newaccesstoken_cmdb" -ForegroundColor Yellow -BackgroundColor Black
$newaccesstokenobject | Export-Csv "$cmdbrefreshtokenpathinfo\$($newaccesstoken_cmdb).csv" -NoTypeInformation
}
}
}
}
$count=$count+1
}while($true)