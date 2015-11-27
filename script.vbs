Dim shell,command

command = "powershell.exe -nologo -command ""C:\Users\windows\Desktop\Play_soundv2.ps1"""

Set shell = CreateObject("WScript.Shell")

shell.Run command,0