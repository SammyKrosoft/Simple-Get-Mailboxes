$FormatEnumerationLimit
$Collection = @()
$Move = $null
#$BatchName = "BatchFinal"
$ExportFilePath = "D:\Scripts\Po\DDLR.csv"

$users = Import-Csv "D:\Scripts\Po\DDL2.csv"
foreach ($list in $users)
{

       #$UserInfo = get-mailbox $id.sAMAccountName
       $DDL = Get-DynamicDistributionGroup $list.Name
      # $DDLAccept = Get-DynamicDistributionGroup $list.Name | select -ExpandProperty @{Name=’AcceptMessageOnlyFrom’;Expression={[System.String]::join(";", ($_.AcceptMessagesOnlyFrom))}} 

     
$DDLAccept = Get-DynamicDistributionGroup $list.Name | SELECT @{Name='AcceptMessageOnlyFrom';Expression={$_.AcceptMessagesOnlyFrom -join ";"}}



       $DDLAcceptSender = Get-DynamicDistributionGroup $list.Name | select @{Name=’AcceptMessageOnlyFromSender’;Expression={[string]::join(";", ($_.AcceptMessagesOnlyFromSendersOrMembers))}} 
       $DDLAcceptDL = Get-DynamicDistributionGroup $list.Name | select @{Name=’AcceptMessageOnlyFromDL’;Expression={[string]::join(";", ($_.AcceptMessagesOnlyFromDLMembers))}}
     


      #$DDLCount = (Get-Recipient -ResultSize unlimited -RecipientPreviewFilter ($DDL.RecipientFilter)).count

    #(Get-Recipient -ResultSize unlimited -RecipientPreviewFilter ($FTE.RecipientFilter)).count

    

     $CurrentCustomObject = [PSCustomObject]@{
                "DDL Name" = $DDL.Name
                "Primary Email Address" = $DDL.PrimarySmtpAddress
                "# of users" = $DDLCount
                #"Accept Message Only From" = [string]::$ddl.AcceptMessagesOnlyFrom
                #"Accept Message Only From" = @{expression={$_.AcceptMessagesOnlyFrom -join ";"}}
                "Accept Message Only From" = $DDLAccept
                
                #@{Name=’GrantSendOnBehalfTo’;Expression={[string]::join(";", ($_.GrantSendOnBehalfTo))}}
                "Accept Message Only From Senders/Memebers" = $DDLAcceptSender
                "Accept Message Only From DL" = $DDLAcceptDL

                #"Accept Message Only From Senders/Memebers" = $DDL.AcceptMessagesOnlyFromSendersOrMembers
                #"Accept Message Only From DL" = $DDL.AcceptMessagesOnlyFromDLMembers
                "Managed By" = $DDL.ManagedBy
                "Recipient Container" = $DDL.RecipientContainer
                
                }
             $Collection += $CurrentCustomObject
           }


$Collection | Export-CSV $ExportFilePath -NoTypeInformation -Encoding UTF8

#Notepad $ExportFilePath  
