<#
  Creates a new rule to move messages from a specific user to a folder (folder at root of mailbox)
  Set the params as required. Admin user needs to be an Exchange Administrator
  Requires ExchangeOnline PowerShell module
#>
#region params
#user to connect to EXO with (Exchange admin)
$adminUser = ''
#$mailboxToProcess= @(@{UserPrincipalName = 'user@domain.com'})
#if using a csv; expects it to contain a list of UPNs with UserPrincipalName column header
$mailboxToProcess= import-csv 'C:\temp\users.csv'
#the folder name to move messages to
$newFolderName = ''
#name to give to the rule
$calRuleName = ''
#service account sending the invites
$serviceAccount = "serviceaccount@domain.com"
$failedUsers = @()
#endregion params

#region intialise stuff
Connect-ExchangeOnline -UserPrincipalName $adminUser
$failedUsers = @()
#endregion initialise stuff

foreach ($m in $mailboxToProcess) {
    $mailbox = $m.UserPrincipalName
    try {
        #Get the mailbox, deal with ' in upns by replacing with ''
        $mbx = Get-Mailbox -filter "userprincipalname -eq '$($mailbox.Replace("'","''"))'" -ErrorAction Stop
        #add full access permissions to the admin account (requried to do the rules even though exchange administrator)
        $null = Add-MailboxPermission -User $adminUser -AccessRights fullaccess -InheritanceType all -Identity $mbx -ErrorAction Stop
        $calRules = Get-InboxRule -Mailbox $mbx | ? {$_.name -eq $calRuleName}
        $calRuleCount = ($calRules | Measure-Object).count
        if ($calRuleCount -gt 0)
        {
            Write-Warning "[$mailbox] Already contains rules names $calRuleName : $calRuleCount"
        }
        else {
            Write-Host "[$mailbox] Adding rule"
            New-InboxRule -Name $calRuleName -Mailbox $mbx -MoveToFolder "$($mailbox):\$($newFolderName)" -From $serviceAccount -MarkAsRead $true -StopProcessingRules $True
        }
        #remove the full access permissions of the admin account
        Remove-MailboxPermission -User $adminUser -AccessRights fullaccess -InheritanceType all -Identity $mbx -ErrorAction Stop -Confirm:$false
    }
    catch {
        Write-Warning "[$mailbox] Something went wrong"
        $failedUsers += $mailbox
    }
}
$failedUsers | ft -AutoSize
$failedUsers | Export-csv c:\temp\failed.csv -NoTypeInformation
