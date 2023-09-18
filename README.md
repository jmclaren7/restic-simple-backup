# Restic SimpleBackup
This AutoIT script heps with the depoyment, configuration and operation of Restic backup and Restic-Browser
1. Copy SimpleBackup.exe to your system, making a folder if desired. You Can alo rename SimpleBackup.exe to whatever you prefer. eg: C:\Program Files\AcmeSimpleBackup\AcmeSimpleBackup.exe
2. Run SimpleBackup.exe and configured the backup path, bucket info and restic password. (For B2, use the AWS variables)
3. Hit apply to save the configuration
4. Run a restic backup from the drop down menu and wait for your first backup to complete successfully
5. From the tools menu, create a scheduled task so the backup will be automatic, customize the task as needed in the Windows Scheduled Task configuration
6. From the tools menu, open Restic Browser to explore and test your backup
