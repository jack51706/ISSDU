::© 2017 AO Kaspersky Lab. All Rights Reserved.


set data_share="\\127.0.0.1\data_share"
net use y: %data_share%
mkdir y:\%COMPUTERNAME%_report
set dp=y:\%COMPUTERNAME%_report
echo %date% %time% %COMPUTERNAME% > %dp%\report.log
fls.exe -lpr \\.\c: >> %dp%\fls.log
::Get Windows reg files
findstr /i "windows\/system32\/config\/system " %dp%\fls.log | findstr /vi "profile" | findstr /vi log | cut -f2 -d" " | cut -f1 -d":" > %dp%\system.reg.inode
for /f "tokens=1" %%a in (%dp%\system.reg.inode) do icat \\.\c: %%a > %dp%\system.reg
findstr /i "windows\/system32\/config\/system " %dp%\fls.log | findstr /vi "profile" | findstr /vi log | cut -f2 -d" " | cut -f1 -d":" > %dp%\software.reg.inode
for /f "tokens=1" %%a in (%dp%\software.reg.inode) do icat \\.\c: %%a > %dp%\software.reg
::Convert reg files
reglookup.exe %dp%\system.reg > %dp%\system.reg.log 
reglookup.exe %dp%\software.reg > %dp%\\software.reg.log
::Get strings from reg files
strings -afel %dp%\system.reg > %dp%\system.str.log
strings -afeb %dp%\system.reg >> %dp%\system.str.log
strings -afel %dp%\software.reg > %dp%\software.str.log
strings -afeb %dp%\software.reg >> %dp%\software.str.log
::Get Logs
grep -i "windows\/system32\/winevt/logs/system.evtx" %dp%\fls.log | cut -f2 -d" " | cut -f1 -d":" > %dp%\system.evtx.inode
for /f "tokens=1" %%a in (%dp%\system.evtx.inode) do icat \\.\c: %%a > %dp%\system.evtx
findstr /i "windows\/system32\/winevt/logs/security.evtx" %dp%\fls.log | cut -f2 -d" " | cut -f1 -d":" > %dp%\security.evtx.inode
for /f "tokens=1" %%a in (%dp%\security.evtx.inode) do icat \\.\c: %%a > %dp%\security.evtx
strings -afeb %dp%\system.evtx > %dp%\system.evtx.str.log
strings -afel %dp%\system.evtx >> %dp%\system.evtx.str.log
strings -afeb %dp%\security.evtx > %dp%\security.evtx.str.log
strings -afel %dp%\security.evtx >> %dp%\security.evtx.str.log
::Conv evtx
evtxexport.exe %dp%\system.evtx > %dp%\system.evtx.res.log
::get evt logs 
findstr /i "windows\/system32\/config/SysEvent.Evt" %dp%\fls.log | cut -f2 -d" " | cut -f1 -d":" > %dp%\SysEvent.Evt.inode
for /f "tokens=1" %%a in (%dp%\SysEvent.Evt.inode) do icat \\.\c: %%a > %dp%\SysEvent.Evt
findstr /i "windows\/system32\/config/SecEvent.Evt" %dp%\fls.log | cut -f2 -d" " | cut -f1 -d":" > %dp%\SecEvent.Evt.inode
for /f "tokens=1" %%a in (%dp%\SecEvent.Evt.inode) do icat \\.\c: %%a > %dp%\SecEvent.Evt
::get strings from evt
strings -afeb %dp%\SysEvent.Evt > %dp%\SysEvent.Evt.str.log
strings -afel %dp%\SysEvent.Evt >> %dp%\SysEvent.Evt.str.log
strings -afeb %dp%\SecEvent.Evt > %dp%\SecEvent.Evt.str.log
strings -afel %dp%\SecEvent.Evt >> %dp%\SecEvent.Evt.str.log
::Conv evt
evtexport.exe %dp%\SysEvent.Evt > %dp%\SysEvent.Evt.res.log
::Do some search 8) 
findstr /i "powershell" %dp%\*.log >> %dp%\report.log
