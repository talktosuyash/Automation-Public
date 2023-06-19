#############################################################################################
#   Script to Generate Access and Refresh Token first time using Zoho clientid and secretid #
#                                                                                           #
#                                                                                           #
#############################################################################################

$mainpath1 = ".\manageengin"
$checkfolderpresent_important1 = Get-ChildItem $mainpath
if(($checkfolderpresent_important1 | where { $_.Name -contains "Important" }).Name -ne "Important" ){
try{
New-Item -Path '.\manageengin\zohoinfo' -ItemType Directory -ErrorAction SilentlyContinue
Write-Host "Zohoinfo Folder Created into '.\manageengin\'path please add zohodetails.csv file with following headers `n clientId,clientSecret,authCode-AllRequests,authCode-AllAssets,authCode-ALLCMDB with respective values" -ForegroundColor Yellow -BackgroundColor Black
break
}
catch{

Write-Host "zohoinfo Folder Already Present"
}
}

if(($checkfolderpresent_important1 | where { $_.Name -contains "Token CSVs" }).Name -ne "Token CSVs"){
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
$response,$clientId, $clientSecret, $authorizationCode, $body, $response, $accessToken, $refreshToken, $csvpath,$authorizationCode_req, $zohoappdetailspath, $tokenobject, $presentdate = $null
$assetcsvpath = ".\manageengin\Token CSVs\donotdelete"
$requestcsvpath = ".\manageengin\Token CSVs\donotdelete"
$CMDBcsvpath = ".\manageengin\Token CSVs\donotdelete"
$zohoappdetailspath = ".\manageengin\Important\zohoinfo" 
$zohoappdetails = import-csv "$zohoappdetailspath\zohodetails.csv"
$accesstokencsv = Get-ChildItem $assetcsvpath
$accessrequesttokencsv = Get-ChildItem $requestcsvpath
$accesscmdbtokencsv = Get-ChildItem $CMDBcsvpath
$clientId = $zohoappdetails.clientId
$clientSecret = $zohoappdetails.clientSecret
#$redirectUri = "https://your.redirect.uri"
$authorizationCode = $zohoappdetails.'authCode-AllAssets'
$authorizationCode_req = $zohoappdetails.'authCode-AllRequests'
$authorizationCode_CMDB = $zohoappdetails.'authCode-ALLCMDB'

#################### Asset ######################################
if($authorizationCode -ne $null){
$body = @{
    grant_type = "authorization_code"
    client_id = $clientId
    client_secret = $clientSecret
    code = $authorizationCode
}

$headers = @{
    "Content-Type" = "application/x-www-form-urlencoded"
}

$presentdate = Get-Date -Format 'dd/MM/yyyy hh:mm tt'

$response = Invoke-RestMethod -Method POST -Uri "https://accounts.zoho.in/oauth/v2/token" -Headers $headers -Body $body


$accessToken = $response.access_token
$refreshToken = $response.refresh_token

Write-Host "Firsttime Access Token : $accessToken" -ForegroundColor White -BackgroundColor Black
$tokenobject = New-Object -TypeName psobject
$tokenobject | Add-Member -MemberType NoteProperty -Name accesstoken -Value $accessToken
$tokenobject | Add-Member -MemberType NoteProperty -Name refreshtoken -Value $refreshToken
$tokenobject | Add-Member -MemberType NoteProperty -Name time -Value $presentdate
if($accesstokencsv -ne $null -and $accesstokencsv.BaseName -ne $accessToken)
{
Remove-Item $accesstokencsv.pspath -Verbose

$tokenobject | export-csv "$assetcsvpath\$($accessToken)_asset.csv" -NoTypeInformation

Start-Sleep -Seconds 5 
}
else{

if($accesstokencsv -eq $null)
{
$tokenobject | export-csv "$assetcsvpath\$($accessToken)_asset.csv" -NoTypeInformation

}
if( $assetcsvpath.name -eq $accessToken)
{
continue
}
}
}
##################### Requests ###################################
if($authorizationCode_req -ne $null){
$headers,$presentdate,$response, $accessToken, $refreshToken, $tokenobject, $accessToken, $refreshToken = $null
$body = @{
    grant_type = "authorization_code"
    client_id = $clientId
    client_secret = $clientSecret
    code = $authorizationCode_req
}

$headers = @{
    "Content-Type" = "application/x-www-form-urlencoded"
}

$presentdate = Get-Date -Format 'dd/MM/yyyy hh:mm tt'

$response = Invoke-RestMethod -Method POST -Uri "https://accounts.zoho.in/oauth/v2/token" -Headers $headers -Body $body


$accessToken = $response.access_token
$refreshToken = $response.refresh_token

Write-Host "Firsttime Access Token : $accessToken" -ForegroundColor White -BackgroundColor Black
$tokenobject = New-Object -TypeName psobject
$tokenobject | Add-Member -MemberType NoteProperty -Name accesstoken -Value $accessToken
$tokenobject | Add-Member -MemberType NoteProperty -Name refreshtoken -Value $refreshToken
$tokenobject | Add-Member -MemberType NoteProperty -Name time -Value $presentdate
if($accessrequesttokencsv -ne $null -and $accessrequesttokencsv.BaseName -ne $accessToken)
{
Remove-Item $accessrequesttokencsv.pspath -Verbose

$tokenobject | export-csv "$requestcsvpath\$($accessToken)_request.csv" -NoTypeInformation

Start-Sleep -Seconds 5 
}
else{

if($accessrequesttokencsv -eq $null)
{
$tokenobject | export-csv "$requestcsvpath\$($accessToken)_request.csv" -NoTypeInformation

}
if( $requestcsvpath.name -eq $accessToken)
{
continue
}
}
}

##################### CMDB ###################################
if($authorizationCode_CMDB -ne $null){
$headers,$presentdate,$response, $accessToken, $refreshToken, $tokenobject, $accessToken, $refreshToken = $null
$body = @{
    grant_type = "authorization_code"
    client_id = $clientId
    client_secret = $clientSecret
    code = $authorizationCode_CMDB
}

$headers = @{
    "Content-Type" = "application/x-www-form-urlencoded"
}

$presentdate = Get-Date -Format 'dd/MM/yyyy hh:mm tt'

$response = Invoke-RestMethod -Method POST -Uri "https://accounts.zoho.in/oauth/v2/token" -Headers $headers -Body $body


$accessToken = $response.access_token
$refreshToken = $response.refresh_token

Write-Host "Firsttime Access Token : $accessToken" -ForegroundColor White -BackgroundColor Black
$tokenobject = New-Object -TypeName psobject
$tokenobject | Add-Member -MemberType NoteProperty -Name accesstoken -Value $accessToken
$tokenobject | Add-Member -MemberType NoteProperty -Name refreshtoken -Value $refreshToken
$tokenobject | Add-Member -MemberType NoteProperty -Name time -Value $presentdate
if($accesscmdbtokencsv -ne $null -and $accesscmdbtokencsv.BaseName -ne $accessToken)
{
Remove-Item $accesscmdbtokencsv.pspath -Verbose

$tokenobject | export-csv "$CMDBcsvpath\$($accessToken)_cmdb.csv" -NoTypeInformation

Start-Sleep -Seconds 5 
}
else{

if($accesscmdbtokencsv -eq $null)
{
$tokenobject | export-csv "$CMDBcsvpath\$($accessToken)_cmdb.csv" -NoTypeInformation

}
if( $requestcsvpath.name -eq $accessToken)
{
continue
}
}
}
}
