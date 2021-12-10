#Collect onpremise Exchange Mailbox stats
#Author - Antonio Rodrigues - antonio.rodrigues@ssc-spc.gc.ca
#Last Updated - Dec 4, 2021

#Original location :
#$OutputFile = "C:\exchange-monitoring\reporting\pickup_folder\mb.csv"
#Sam Lab location:
$OutputFile = "C:\Temp\pickup_folder\mb.csv"

# NOTE: You can omit the below if you run the script from an Exchange Management Shell:
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

$Databases = Get-MailboxDatabase

$Dbprogresscounter = 0
$DatabasesCount = $Databases.Count
$ObjectCollectionToExport = @()
Foreach ($database in $Databases) {
    write-progress -Activity "Parsing databases" -Status "Now in database $($database.Name) ..." -PercentComplete $($Dbprogresscounter/$DatabasesCount*100)
    $Mailboxes = $null
    $Mailboxes = Get-Mailbox -ResultSize Unlimited -Database $Database -Filter {RecipientTypeDetails -ne "DiscoveryMailbox"}| Select Name,PrimarySMTPAddress,REcipientTypeDetails,RecipientType, LitigationHoldEnabled, IssueWarningQuota, ProhibitSendQuota, ProhibitSendReceiveQuota, RetainDeletedItemsFor, UseDatabaseQuotaDefaults, SingleItemRecoveryEnabled, RecoverableItemsQuota, UseDatabaseRetentionDefaults, Database
    Write-Host "Found $($Mailboxes.count) mailboxes on database $($Database.name) ..." -ForegroundColor Green

    #region Inserting ESDC mailbox info collection Routine from Antonio Rodriguez script ###########

    Foreach ($mbx in $Mailboxes) {
        $stats = Get-MailboxStatistics $mbx.name | Select-Object Lastlogontime, TotalItemSize, Itemcount, TotalDeletedItemSize
        $user = Get-User $_.Name | Select-Object SID

        $Object = New-Object -TypeName PSObject -Property @{
            RecipientType = $mbx.RecipientType
            LitigationHoldEnabled = $mbx.LitigationHoldEnabled
            IssueWarningQuota = $mbx.IssueWarningQuota
            ProhibitSendQuota = $mbx.ProhibitSendQuota
            ProhibitSendReceiveQuota = $mbx.ProhibitSendReceiveQuota
            RetainDeletedItemsFor = $mbx.RetainDeletedItemsFor
            UseDatabaseQuotaDefaults = $mbx.UseDatabaseQuotaDefaults
            SingleItemRecoveryEnabled = $mbx.SingleItemRecoveryEnabled
            RecoverableItemsQuota = $mbx.RecoverableItemsQuota
            UseDatabaseRetentionDefaults = $mbx.UseDatabaseRetentionDefaults
            Database = $mbx.Database
            Lastlogontime = $stats.Lastlogontime
            TotalItemSize = $stats.TotalItemSize
            Itemcount = $stats.Itemcount
            TotalDeletedItemSize = $stats.TotalDeletedItemSize
            SID = $user.SID
        }
        $ObjectCollectionToExport += $Object
        $Dbprogresscounter++
    }
    #endregion End of Antonio Routine
}

#Now dumping all the information from the $ObjectCollectionToExport variable into the file
$ObjectCollectionToExport | Export-Csv $OutputFile -NoTypeInformation -Encoding 'UTF8'

#Now appending string to the output file to indicate the script finished
Add-Content -Path $OutputFile -Value "`n#####Script completed successfully######"
