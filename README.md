# Automate-Creating-Outlook-Rules
PowerShell script that gathers a set of Outlook Rules and applies them to a different User.


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



## Table of Contents 

1. [Running Locally](#running-locally)







## Running Locally
1. Clone this repo
1. 
1. 
1. 


### Run the Front-End
