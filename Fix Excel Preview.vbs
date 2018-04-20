' This closes Excel, then launches a hidden Excel window.
' This hidden window will be automatically used when trying to "preview" an Excel file, because Excel
' tries to use the first Excel that launched to display previews.
' This means the preview will work, because the hidden window won't have an Open Dialog Box that blocks Excel from opening a file
' Run this on computer startup, or if you need to restart Excel, just run this script to close all Excel instances 
' and start a ghost one to serve as the "Preview Handler".

Set oShell = WScript.CreateObject("WScript.Shell")
oShell.Run("taskkill /f /im Excel.exe"),0,True ' Close all currently running Windows so we know the next line will launch the first instance
oShell.Run("Excel.exe"),0,True ' Launch a hidden Excel window to serve as the Preview Handler
