<#
  Use EWS to create a folder in the mailbox @ the mailbox root (i.e. not a subfolder of Inbox
  Uses an app registration with certificate for auth and required permissions
  Set the params in the params region
  Should add some logging
#>
 
#region params
#Azure tenant ID
$tenantID = ''
#App registration with exchange permissions
$clientID = ''
#path to EWS dll 
$ewsPath = "C:\temp\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
#Cert thumbnail for authentication using the app 
$cert = ''
#Name of the folder to create
$newFolderName = ''
 #csv to import; column header PrimarySMTPAddress
$csvPath = ''
#endregion params

#region initialise stuff
$mailboxToProcess = import-csv $csvPath
$outlookScopes = https://outlook.office365.com/.default
Add-Type -Path $ewsPath
$ClientCert = Get-Item "Cert:\CurrentUser\My\$($cert)"
$failedUsers = @()
#endregion initialise stuff
 
#region process mailbox
foreach ($mbx in $mailboxToProcess) {
    $mailbox = $mbx.PrimarySMTPAddress
    Add-LogMessage $LogDebug "[$mailbox] Connecting to mailbox" 
    #region setup ews connection
    $ewsToken = get-msaltoken -ClientId $clientID -TenantId $TenantId -Scopes $outlookScopes -RedirectUri 'http://localhost' -ClientCertificate $clientCert
    $ews = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList 'Exchange2010_SP2'
    $ews.Url = 'https://outlook.office365.com/EWS/Exchange.asmx'
    $ews.UseDefaultCredentials = $false
    $ews.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$ewstoken.AccessToken
    #$ews.traceenabled = "true" #uncomment for troubleshooting
    $ews.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $mailbox)
    #https://techcommunity.microsoft.com/t5/exchange/ews-with-oauth-quot-an-internal-server-error-occurred-the/m-p/3609047
    $ews.HttpHeaders.Add("X-AnchorMailbox", $mailbox)
    #endregion setup ews connection
    #region create the new folder
    #bind to the Mailbox root folder; added retry with delay to get around some issues (possibly throttling) 
    $mbRootFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
    try {
        $mbFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ews, $mbRootFolderId)
    }
    catch {
        #add a short delay and try 1 more time
        Start-Sleep 60
        $ewsToken = get-msaltoken -ClientId $clientID -TenantId $TenantId -Scopes $outlookScopes -RedirectUri 'http://localhost' -ClientCertificate $clientCert
        $ews.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$ewstoken.AccessToken
        $ews.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $mailbox)
        try {
            #bind to the Mailbox root folder
            $mbFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ews, $mbRootFolderId)
        }
        catch {
            $failedUsers += $mailbox
            continue
        }
    }
    #Setup the new folder
    $NewFolder = new-object Microsoft.Exchange.WebServices.Data.Folder($ews)  
    $NewFolder.DisplayName = $newFolderName
    $NewFolder.FolderClass = "IPF.Note"
    #Check if the folder already exists
    #Define folder veiw, really only want to return one object  
    $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)  
    #Define a search folder; searcheson the DisplayName of the folder  
    $sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$newFolderName)  
    #Do the search  
    try {
        $findFolderResults = $ews.FindFolders($mbFolder.Id,$sfSearchFilter,$fvFolderView)  
    }
    catch {
        Start-sleep 10
        try {
            $findFolderResults = $ews.FindFolders($mbFolder.Id,$sfSearchFilter,$fvFolderView)  
        }
        catch {
            $failedUsers += $mailbox
            continue
        }
    }
    if ($findFolderResults.TotalCount -ne 0){  
        Write-Warning "[$mailbox] Folder already exists"
    }
    elseif ($findFolderResults.TotalCount -eq 0) {
        Write-Output "[$mailbox] Creating folder" 
        try {
            $newFolder.Save($mbFolder.Id)  
        }
        catch {
            Write-Warning "[$mailbox] Failed to create folder"
            $failedUsers += $mailbox
            continue
        }
    } 
    else{  
        Write-Warning "[$mailbox] Something odd happened!!"
    }
    #endregion create the new folder
}
#endregion process mailbox
$failedUsers | ft
$failedUsers | export-csv C:\temp\failedUsers.csv -NoTypeInformation
