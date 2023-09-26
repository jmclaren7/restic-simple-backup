# Restic SimpleBackup
This AutoIT script helps with the configuration and operation of Restic backup. Some parameters are simply passed to Restic, you can find more information about Restic [here](https://restic.readthedocs.io/en/stable/).

<p align="center">
  <img src="https://github.com/jmclaren7/restic-simple-backup/blob/main/extra/main-window.jpg?raw=true">
</p>

## Features
* Configuration GUI
* Single executable - Integrates Restic.exe
* Integrated Restic-Browser to explore and restore from backups
* Uses the same enviromental variables from Restic
* Create a scheduled task from GUI
* Multiple backup profiles for seperate backup tasks
* Send email with backup status

## SimpleBackup Setup
1. Download the latest SimpleBackup.exe from the releases page
2. Copy SimpleBackup.exe to your system, making a folder if desired. You can also rename SimpleBackup.exe to whatever you prefer. (A configuration file that matches the executable name is later generated)
4. Run SimpleBackup.exe and configure the SimpleBackup and Restic options
5. Click apply to save the configuration
6. Initialize your Restic repository using the command box (init)
7. Start a Restic backup from the command box and wait for your first backup to complete successfully
8. From the tools menu, create a scheduled task so the backup will be automatic
9. From the tools menu, open Restic Browser to explore and verify your backup

## BackBlaze B2 Setup
To add resistance to ransomware you'll need to use the b2 cli so you can specify exact permissions, b2.exe can be downloaded from their GitHub. The B2 web interface has many other limitations so being familiar with the cli is highly recommended.

1. Create a B2 bucket (web or cli) and save the bucket name and s3 endpoint in a safe location
2. Set bucket lifecycle rules (web or cli) to delete versions after a specified number of days (I recommend at least 90 days).
3. Using the B2 cli, create a key to access the bucket with and save the id and key to a safe location (the key begins with the letter K)
```shell
b2.exe create-key --bucket real-bucket-name any-key-name listBuckets,readFiles,writeFiles,listFiles
```
5. Use the S3 compatible endpoint provided by Backblaze. In SimpleBackup, set the AWS_* variables and set RESTIC_REPOSITORY using this format: s3:\<end-point\>/\<bucket-name\> 

## Example Minimal Config
```ini
Setup_Password=1234
Backup_Path=C:\Data
Backup_Prune=--keep-last 7 --keep-daily 7 --keep-weekly 52
RESTIC_PASSWORD=pppppppppppppppppppppppppppppp
RESTIC_REPOSITORY=s3:s3.us-west-001.backblazeb2.com/real-bucket-name
AWS_ACCESS_KEY_ID=iiiiiiiiiiiiiiiiiiiiiiiii
AWS_SECRET_ACCESS_KEY=kkkkkkkkkkkkkkkkkkkkkkkkkkkkkkk
```


## Configuration Options
Besides the enviromental variables available to use with Restic, SimpleBackup offers the following settings
Setup_Password
Backup_Path
Backup_Prune
Backup_Name
SMTP_Server
SMTP_UserName
SMTP_Password
SMTP_FromAddress
SMTP_FromName
SMTP_ToAddress
SMTP_SendOnFailure (True/False)
SMTP_SendOnSuccess (True/False)
