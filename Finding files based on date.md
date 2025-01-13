# Finding files based on date

```Powershell
Get-ChildItem -Path "C:\YourDirectory" -Recurse | Where-Object { $_.CreationTime -gt [datetime]::Parse("2023-01-01") }

Get-ChildItem -Path "C:\YourDirectory" -Recurse | Where-Object { $_.LastWriteTime -gt [datetime]::Parse("2023-01-01") }

```
From Copilot

Give me a powershell command to search a tree of directories for files created after a specified date<br>How would this change to look for files updated since the specified date?

a{text-decoration:none;color:#464feb}

Sure! Here's a PowerShell command to search a tree of directories for files created after a specified date:

```powershell
Get-ChildItem -Path "C:\YourDirectory" -Recurse | Where-Object { $_.CreationTime -gt [datetime]::Parse("2023-01-01") }
```

To modify this command to search for files updated (modified) since the specified date, you can use the `LastWriteTime` property instead of `CreationTime`:

```powershell
Get-ChildItem -Path "C:\YourDirectory" -Recurse | Where-Object { $_.LastWriteTime -gt [datetime]::Parse("2023-01-01") }
```

Feel free to adjust the path and date as needed. If you have any other questions or need further assistance, just let me know!

Add `-Filter *.xlsx` before the pipe to get just excel files 
