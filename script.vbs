Dim WshShell, strCommand, objHC 

set WshShell = CreateObject("wscript.Shell") 
 

strCommand = "C:\Program Files (x86)\VMware\VMware Horizon View Client\vmware-view.exe" 

Set objHC = WshShell.Exec(strCommand) 

 

' Wait few seconds to let Horizon refresh the application shortcuts list in Start Menu 

wscript.sleep 5000 


' Logoff Horizon Agent 

strCommand = "C:\Program Files (x86)\VMware\VMware Horizon View Client\vmware-view.exe -shutdown" 

Set objHC = WshShell.Exec(strCommand)
 

'Clean up and exit 

Set objHC = Nothing 

Set WshShell = Nothing 

WScript.Quit
