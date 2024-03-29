<#
    
Automate Creating Outlook Rules 1.0:

    Make sure to run this script in administration. 

Required Software:

    PowerShell v7 or higher
Keep in mind this only works on Windows computer. If you want it to work for Mac, you need to choose the particular location To save that CSV file.
    
Things that are currently not working:
   Create a folder in Outlook 
   Add a Rule to move emails to pacific folder 

 
Below list of rules that are currently not working: 
   ReceivedBeforeDate
   ReceivedAfterDate
   MessageTypeMatches
   DeleteSystemCategory
   SendTextMessageNotificationTo
   ExceptIfMessageTypeMatches
   ReceivedAfterDate
   ExceptIfReceivedAfterDate
   ExceptIfReceivedBeforeDate
   ExceptIfFromSubscription
   ExceptIfWithinSizeRangeMinimum
   WithinSizeRangeMaximum
   WithinSizeRangeMinimum  

The script will continue to add rules, it just won't add these particular ones It should not error out. 

Tips:
    - Remember If you're copying So what I have CSV file without editing it.It will copy everything, including follow to 
	Run this command to update 
   Install-Module -Name ExchangeOnlineManagement -Force 
   I recommend after updating to completely close out of PowerShell and reopen a new Terminal. 
 
TODO:
   - Next update look into add: 
      if($_.WithImportance){$Filtered_rules.WithImportance=$_.WithImportance} *****-*****
      if($_.ExceptIfWithImportance){$Filtered_rules.ExceptIfWithImportance=$_.ExceptIfWithImportance} *****-*****
      if($_.MarkImportance){$Filtered_rules.MarkImportance=$_.MarkImportance} *****-*****
      if($_.WithSensitivity){$Filtered_rules.WithSensitivity=$_.WithSensitivity}
      if($_.ExceptIfWithSensitivity){$Filtered_rules.ExceptIfWithSensitivity=$_.ExceptIfWithSensitivity}





#>

#Manually save CSV Mac users: 
#Make sure to Delete the function save file.

#Add this instead: 
# $files = "Your file path choice."


Import-Module ExchangeOnlineManagement

#This command will connect you to exchange and also ask for your credentials.
Connect-ExchangeOnline 

#This commandlet saves CSV file to a particular location of your choice.
function Save-File([string] $initialDirectory ) 

{

[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

$OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$OpenFileDialog.initialDirectory = $initialDirectory
$OpenFileDialog.FileName = "Temporary files"
$OpenFileDialog.filter = "CSV (*.csv)| *.csv"
$OpenFileDialog.ShowDialog() |  Out-Null

return $OpenFileDialog.FileName  

} 

$files = Save-File


$csv= Import-Csv $files

Clear-Host
Write-Host "-----------------------------------------";
   Write-Host -BackgroundColor Magenta "         Automate Creating Outlook Rules      "; 
   Write-Host -BackgroundColor Blue "                Alpha 1.0               ";
     
   Write-Host "-----------------------------------------";
           Write-Host -Object '          "Gathering Information"       ' -ForegroundColor Yellow 
           Write-Host "`n"
           Write-Host -Object 'Please type in user email Below' -ForegroundColor DarkCyan
        Write-Host     -Object '*******************************'
        $User_email = Read-Host "Email"
         pause 
        Write-Host  'Now processing Please wait.' -ForegroundColor blue

$csv| foreach {
   $Filtered_rules= @{} #hashtable
    if ($_.Name){$Filtered_rules.name=$_.Name}

    if($_.Priority){$Filtered_rules.Priority=$_.Priority}

    if($_.BodyContainsWords){$Filtered_rules.SubjectorBodyContainsWords=$_.BodyContainsWords}

    if($_.ExceptIfBodyContainsWords){$Filtered_rules.ExceptIfBodyContainsWords=$_.ExceptIfBodyContainsWords}

    if($_.FlaggedForAction){$Filtered_rules.FlaggedForAction=$_.FlaggedForAction}

    if($_.ExceptIfFlaggedForAction){$Filtered_rules.ExceptIfFlaggedForAction=$_.ExceptIfFlaggedForAction}

    if($_.FromAddressContainsWords){$Filtered_rules.FromAddressContainsWords=$_.FromAddressContainsWords}

    if($_.ExceptIfFromAddressContainsWords){$Filtered_rules.ExceptIfFromAddressContainsWords=$_.ExceptIfFromAddressContainsWords}

    if($_.From){$Filtered_rules.From=$_.From}
    
    if($_.ExceptIfFrom){$Filtered_rules.ExceptIfFrom=$_.ExceptIfFrom}

    if($_.HasAttachment -match $True){$Filtered_rules.HasAttachment= $True} else {
      $Filtered_rules.HasAttachment= $False
    }

    if($_.ExceptIfHasAttachment -match $True ){$Filtered_rules.ExceptIfHasAttachment= $True} else {
      $Filtered_rules.ExceptIfHasAttachment= $False
    }

    if($_.HasClassification){$Filtered_rules.HasClassification=$_.HasClassification}

    if($_.ExceptIfHasClassification){$Filtered_rules.ExceptIfHasClassification=$_.ExceptIfHasClassification}

    if($_.HeaderContainsWords){$Filtered_rules.HeaderContainsWords=$_.HeaderContainsWords}

    if($_.ExceptIfHeaderContainsWords){$Filtered_rules.ExceptIfHeaderContainsWords=$_.ExceptIfHeaderContainsWords}

    if($_.FromSubscription){$Filtered_rules.FromSubscription=$_.FromSubscription}
    
    if($_.ExceptIfFromSubscription){$Filtered_rules.ExceptIfFromSubscription=$_.ExceptIfFromSubscription}

    if($_.MessageTypeMatches){$Filtered_rules.MessageTypeMatches=$_.MessageTypeMatches}

    if($_.ExceptIfMessageTypeMatches){$Filtered_rules.ExceptIfMessageTypeMatches=$_.ExceptIfMessageTypeMatches}

    if($_.MyNameInCcBox -match $True){$Filtered_rules.MyNameInCcBox=$True}else {
      $Filtered_rules.MyNameInCcBox=$False
    }

    if($_.ExceptIfMyNameInCcBox -match $True ){$Filtered_rules.ExceptIfMyNameInCcBox=$True} else {
      $Filtered_rules.ExceptIfMyNameInCcBox=$False
    }

    
    if($_.MyNameInToBox -match $True ){$Filtered_rules.MyNameInToBox=$_.MyNameInToBox= $True} else {
      $Filtered_rules.MyNameInToBox=$_.MyNameInToBox= $False
    }



    if($_.ExceptIfMyNameInToBox -match $True ){$Filtered_rules.ExceptIfMyNameInToBox= $True} else {
      $Filtered_rules.ExceptIfMyNameInToBox= $False
    }

    if($_.MyNameInToOrCcBox -match $True ){$Filtered_rules.MyNameInToOrCcBox= $True} else {
      $Filtered_rules.MyNameInToOrCcBox= $False
    }

    if($_.ExceptIfMyNameInToOrCcBox -match $True){$Filtered_rules.ExceptIfMyNameInToOrCcBox= $True} else {
      $Filtered_rules.ExceptIfMyNameInToOrCcBox= $False
    }

    if($_.MyNameNotInToBox -match $True){$Filtered_rules.MyNameNotInToBox= $True} else {
      $Filtered_rules.MyNameNotInToBox= $False
    }

    if($_.ExceptIfMyNameNotInToBox -match $True){$Filtered_rules.ExceptIfMyNameNotInToBox= $True} else {
      $Filtered_rules.ExceptIfMyNameNotInToBox= $False
    }

    if($_.ReceivedAfterDate  ){$Filtered_rules.ReceivedAfterDate=$_.ReceivedAfterDate}

    if($_.ExceptIfReceivedAfterDate){$Filtered_rules.ExceptIfReceivedAfterDate=$_.ExceptIfReceivedAfterDate}
    
    if($_.ReceivedBeforeDate){$Filtered_rules.ReceivedBeforeDate=$_.ReceivedBeforeDate}

    if($_.ExceptIfReceivedBeforeDate){$Filtered_rules.ExceptIfReceivedBeforeDate=$_.ExceptIfReceivedBeforeDate}

    if($_.RecipientAddressContainsWords){$Filtered_rules.RecipientAddressContainsWords=$_.RecipientAddressContainsWords}

    if($_.ExceptIfRecipientAddressContainsWords){$Filtered_rules.ExceptIfRecipientAddressContainsWords=$_.ExceptIfRecipientAddressContainsWords}

    if($_.SentOnlyToMe -match $True){$Filtered_rules.SentOnlyToMe= $True} else {
      $Filtered_rules.SentOnlyToMe= $False
    }

    if($_.ExceptIfSentOnlyToMe -match $True){$Filtered_rules.ExceptIfSentOnlyToMe= $True} else {
      $Filtered_rules.ExceptIfSentOnlyToMe= $False
    }
    
    if($_.SentTo){$Filtered_rules.SentTo=$_.SentTo}

    if($_.ExceptIfSentTo){$Filtered_rules.ExceptIfSentTo=$_.ExceptIfSentTo}
    
    if($_.SubjectContainsWords){$Filtered_rules.SubjectContainsWords=$_.SubjectContainsWords}

    if($_.ExceptIfSubjectContainsWords){$Filtered_rules.ExceptIfSubjectContainsWords=$_.ExceptIfSubjectContainsWords}

    if($_.SubjectOrBodyContainsWords){$Filtered_rules.SubjectOrBodyContainsWords=$_.SubjectOrBodyContainsWords}
    
    if($_.EnabledExceptIfSubjectOrBodyContainsWords){$Filtered_rules.EnabledExceptIfSubjectOrBodyContainsWords=$_.EnabledExceptIfSubjectOrBodyContainsWords}

    if($_.WithImportance){$Filtered_rules.WithImportance=$_.WithImportance}
    
    if($_.ExceptIfWithImportance){$Filtered_rules.ExceptIfWithImportance=$_.ExceptIfWithImportance}

    if($_.WithinSizeRangeMaximum){$Filtered_rules.WithinSizeRangeMaximum=$_.WithinSizeRangeMaximum}
    
    if($_.ExceptIfWithinSizeRangeMaximum){$Filtered_rules.ExceptIfWithinSizeRangeMaximum=$_.ExceptIfWithinSizeRangeMaximum}

    if($_.WithinSizeRangeMinimum){$Filtered_rules.WithinSizeRangeMinimum=$_.WithinSizeRangeMinimum}

    if($_.ExceptIfWithinSizeRangeMinimum){$Filtered_rules.ExceptIfWithinSizeRangeMinimum=$_.ExceptIfWithinSizeRangeMinimum}

    if($_.WithSensitivity){$Filtered_rules.WithSensitivity=$_.WithSensitivity}

    if($_.ExceptIfWithSensitivity){$Filtered_rules.ExceptIfWithSensitivity=$_.ExceptIfWithSensitivity}
    
    if($_.ApplyCategory){$Filtered_rules.ApplyCategory=$_.ApplyCategory}

    if($_.ApplySystemCategory){$Filtered_rules.ApplySystemCategory=$_.ApplySystemCategory}

    if($_.CopyToFolder){$Filtered_rules.CopyToFolder=$_.CopyToFolder}

    if($_.DeleteMessage -match $True){$Filtered_rules.DeleteMessage= $True} else {
      $Filtered_rules.DeleteMessage= $False
    }

    if($_.DeleteSystemCategory -match $True){$Filtered_rules.DeleteSystemCategory= $True} else {
      $Filtered_rules.DeleteSystemCategory= $False
    }

    if($_.ForwardAsAttachmentTo){$Filtered_rules.ForwardAsAttachmentTo=$_.ForwardAsAttachmentTo}

    if($_.ForwardTo){$Filtered_rules.ForwardTo=$_.ForwardTo}

    if($_.MarkAsRead -match $True){$Filtered_rules.MarkAsRead= $True} else {
      $Filtered_rules.MarkAsRead= $False
    }

    if($_.MarkImportance){$Filtered_rules.MarkImportance=$_.MarkImportance}

    if($_.PinMessage -match $True ){$Filtered_rules.PinMessage= $True} else {
      $Filtered_rules.PinMessage= $False
    }

    if($_.RedirectTo){$Filtered_rules.RedirectTo=$_.RedirectTo}

    if($_.SendTextMessageNotificationTo){$Filtered_rules.SendTextMessageNotificationTo=$_.SendTextMessageNotificationTo}

    if($_.StopProcessingRules -match $True){$Filtered_rules.StopProcessingRules= $True} else {
      $Filtered_rules.StopProcessingRules= $False
    }





$name.name | ForEach-Object {
 
   ## Define paramters for Send-MailMessage
   $rules= @{
     
      Name                                 = $Filtered_rules.name
      Priority                             = $Filtered_rules.Priority
      BodyContainsWords                    = $Filtered_rules.SubjectorBodyContainsWords
      ExceptIfBodyContainsWords            = $Filtered_rules.ExceptIfBodyContainsWords
      FlaggedForAction                     = $Filtered_rules.FlaggedForAction
      
      ExceptIfFlaggedForAction             = $Filtered_rules.ExceptIfFlaggedForAction
      FromAddressContainsWords             = $Filtered_rules.FromAddressContainsWords
      ExceptIfFromAddressContainsWords     = $Filtered_rules.ExceptIfFromAddressContainsWords 
      From                                 = $Filtered_rules.From
      ExceptIfFrom                         = $Filtered_rules.ExceptIfFrom
      
      HasClassification                    = $Filtered_rules.HasClassification
      ExceptIfHasClassification            = $Filtered_rules.ExceptIfHasClassification
      HeaderContainsWords                  = $Filtered_rules.HeaderContainsWords 
      ExceptIfHeaderContainsWords          = $Filtered_rules.ExceptIfHeaderContainsWords
      RecipientAddressContainsWords        = $Filtered_rules.RecipientAddressContainsWords 
      
      
      SentTo                               = $Filtered_rules.SentTo
      ExceptIfSentTo                       = $Filtered_rules.ExceptIfSentTo
      SubjectContainsWords                 = $Filtered_rules.SubjectContainsWords
      ExceptIfSubjectContainsWords         = $Filtered_rules.ExceptIfSubjectContainsWords
      SubjectOrBodyContainsWords           = $Filtered_rules.SubjectOrBodyContainsWords
      
      MyNameInCcBox                        = $Filtered_rules.MyNameInCcBox
      ExceptIfMyNameInCcBox                = $Filtered_rules.ExceptIfMyNameInCcBox
      MyNameInToBox                        = $Filtered_rules.MyNameInToBox
      ExceptIfMyNameInToBox                = $Filtered_rules.ExceptIfMyNameInToBox
      MyNameInToOrCcBox                    = $Filtered_rules.MyNameInToOrCcBox
      
      ExceptIfMyNameInToOrCcBox            = $Filtered_rules.ExceptIfMyNameInToOrCcBox
      MyNameNotInToBox                     = $Filtered_rules.MyNameNotInToBox
      ExceptIfMyNameNotInToBox             = $Filtered_rules.ExceptIfMyNameNotInToBox
      SentOnlyToMe                         = $Filtered_rules.SentOnlyToMe
      ExceptIfSentOnlyToMe                 = $Filtered_rules.ExceptIfSentOnlyToMe
     
      ApplyCategory                        = $Filtered_rules.ApplyCategory
      ApplySystemCategory                  = $Filtered_rules.ApplySystemCategory
      CopyToFolder                         = $Filtered_rules.CopyToFolder
      DeleteMessage                        = $Filtered_rules.DeleteMessage
      ForwardAsAttachmentTo                = $Filtered_rules.ForwardAsAttachmentTo
      
      MarkAsRead                           = $Filtered_rules.MarkAsRead
      ForwardTo                            = $Filtered_rules.ForwardTo
      PinMessage                           = $Filtered_rules.PinMessage
      RedirectTo                           = $Filtered_rules.RedirectTo
      StopProcessingRules                  = $Filtered_rules.StopProcessingRules

      HasAttachment                        = $Filtered_rules.HasAttachment
      ExceptIfHasAttachment                = $Filtered_rules.ExceptIfHasAttachment
      ExceptIfRecipientAddressContainsWords = $Filtered_rules.ExceptIfRecipientAddressContainsWords
      Mailbox                               = $User_email
      



      
      
  
    

 }

  New-InboxRule @rules


}
}
#Just a default Disconnect closing. I will add more menus later on and will also add a try and catch error.
Write-Host Script is done All rules have been copied and added over now closing. -foregroundcolor red
Disconnect-ExchangeOnline -Confirm:$false

