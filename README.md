# mailboxPermissions: A powershell script to help configure common Exchange Admin Center Tasks

### Features:
* An easier way to configure your common day to day tasks in exchange
* Select your desired configuration by entering a number from a list of tasks that are shown
* Opens Powershell as Administrator
* Connects to Exchange Online
* Allows the session to stay open if you have more tasks to perform
* Allows to add/remove multiple users
* Will remove the session after script completes

### Known limitations:
* Not friendly on going back on your configuration task if you've selected the wrong option
    * There are confirmation questions if a user profile is found or not but not to re-select a configuration tasks.
        * Work around: Hit enter until you're able to get the "Would you like to end the session? (Y/N)" prompt and enter N or No to re-run the script
* Retreiving information
    * This script was created for simple configuration purposes with the idea in mind that we do not need to retreive information.
