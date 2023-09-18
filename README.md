# Restic SimpleBackup
This AutoIT script helps with the deployment, configuration and operation of Restic backup and Restic-Browser
1. Copy SimpleBackup.exe to your system, making a folder if desired. You Can also rename SimpleBackup.exe to whatever you prefer. eg: C:\Program Files\AcmeSimpleBackup\AcmeSimpleBackup.exe
2. Run SimpleBackup.exe and configured the backup path, bucket info and Restic password. (For B2, use the AWS variables)
3. Hit apply to save the configuration
4. Run a Restic backup from the drop down menu and wait for your first backup to complete successfully
5. From the tools menu, create a scheduled task so the backup will be automatic, customize the task as needed in the Windows Scheduled Task configuration
6. From the tools menu, open Restic Browser to explore and test your backup


B2 Ransomware
To add resistance to ransomware you'll need to use the b2 cli so you can specify exact permissions, b2.exe can be downloaded from their GitHub

1. Command Example:
> b2.exe create-key --bucket real-buckit-name any-key-name listBuckets,readFiles,writeFiles,listFiles
2. Copy id and key to a safe location and to the SimpleBackup configuration
3. Set bucket lifecycle rules (web or cli) to delete versions after a specified number of days, I would recommend 90 days.
