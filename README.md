# PSGui
A graphical user interphase for launching powershell commands including asynchronously executed and tracked jobs.

# DESCRIPTION

PSGui is very much what the name implies.
It makes executing and keeping track of all executed commands quite a bit easier,
especially if the user does not have a lot of powershell knowledge.

Select a command, fill in all required data and click execute.
A double click at the finished (running or failed job) shows the output of the job.
Right clicking on a not running job will delete it.

The gui features full keyboard support and can be used without mouse input.
Escape to close, enter to open results, delete to remove jobs.

# IMAGES

http://imgur.com/a/N1tez

# SETUP:
1. Drop all module files (.psm1) with functions that you want to use in the subfolder "PSGui".
2. Edit the text file "commands.txt" and fill in 1 command per line that you want to use in the interface. You can use default functions that powreshell provides (Like "Test-Connection"), or those from your modules.
3. If you want to show the gui without ever having tu open a powershell yourself you can create a shortcut in the same folder with the destination 
    C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -command "Write-Host 'Starting...'; .\PSGui.ps1; exit"
3. ???
4. Profit
