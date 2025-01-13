# Finding files based on date

```Powershell

Get-ChildItem -Path "C:\YourDirectory" -Recurse | Where-Object { $_.CreationTime -gt [datetime]::Parse("2023-01-01") }

Get-ChildItem -Path "C:\YourDirectory" -Recurse | Where-Object { $_.LastWriteTime -gt [datetime]::Parse("2023-01-01") }

```

