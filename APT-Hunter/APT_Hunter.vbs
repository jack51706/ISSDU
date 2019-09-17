'APT_Hunter.vbs 2.0
'Author: @MrRed_Panda
'Copyright (c) 2016, Hao Wang
' Used to hunt for Advanced Persistent Threat (APT)  in the Windows environment
' Require privileged credentials such as domain administrator account
' Require at least Microsoft .net 4.6. 

'++++++++Input domain admin username and password below++++++++++++
On Error Resume Next

username = ""
password = "" 

'++++++++Input domain admin username and password above++++++++++++
'++++++++Define global variables below:++++++++++++

Const ForWriting=2, ForAppending = 8, ForReading = 1
Const HKEY_LOCAL_MACHINE = &H80000002
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = Wscript.ScriptFullName
Set objFile = FSO.GetFile(strPath)
strFolder = FSO.GetParentFolderName(objFile) 
inputfile = strFolder & "\servers.txt"

'Command Line Help Page############################################################################
Set objArgs = WScript.Arguments
intArgCount = objArgs.Count 

If  LCase(objArgs.Item(0)) = "-h" Or LCase(objArgs.Item(0)) = "--help" Or objArgs.Item(0) = " " Then
	WScript.Echo 
	WScript.Echo "###############################################################################################################################################################################################################################################"
	WScript.Echo "Welcome to use APT Hunter 2.0"& vbNewLine
	WScript.Echo "Most organizations are left speechless as 90% of all intrusions are now discovered due to 3rd party notification. And in many cases, the APT has been on your network for years."& vbNewLine
	WScript.Echo "Our goal is to hunt the APT."& vbNewLine
	WScript.Echo " [USAGE]: APT_Hunter.vbs  [OPTIONS]" & vbNewLine
	WScript.Echo vbTab & "cscript APT_Hunter.vbs  /c  username passsword  targetlists" & vbNewLine
	WScript.Echo vbTab & "cscript APT_Hunter.vbs  /dump" & vbNewLine
	WScript.Echo " [OPTIONS]:" & vbNewLine
	WScript.Echo vbTab & vbTab  & "/c"&vbTab&vbTab&vbTab&vbTab&"command line mode, please pass the arguments as username->password->target lists file" & vbNewLine
    WScript.Echo vbTab & vbTab  & "/dump"&vbTab&vbTab&vbTab&vbTab&"none command line mode, please update the corresponding creds variables from the script itself and include all the IPs where you want to check into the servers.txt file under the root folder " & vbNewLine
	WScript.Echo vbTab & vbTab  & "/v"&vbTab&vbTab&vbTab&vbTab&"current version of the tool" & vbNewLine
	WScript.Echo "################################################################################################################################################################################################################################################"
	WScript.Quit()
End If

If LCase(objArgs.Item(0)) <> "/c" AND LCase(objArgs.Item(0)) <> "/dump"  AND  LCase(objArgs.Item(0)) <> "/v" Then 
	WScript.Echo "Syntax error ! Please refer to the help page by typing [cscript APT_Hunter.vbs  -h] or [cscript APT_Hunter.vbs  --help]  "
	WScript.Quit
End If

If  LCase(objArgs.Item(0)) = "/dump" AND intArgCount = 1Then 

WScript.Echo "Starting none comand line mode" &vbNewLine
WScript.Echo " Please make sure to update the corresponding creds variables from the script before use this mode" &vbNewLine
WScript.Echo " Please make sure to include all the IPs where you want to check into the servers.txt file under the root folder before use this mode" &vbNewLine

elseIf  LCase(objArgs.Item(0)) = "/v" AND intArgCount = 1Then 
WScript.Echo "APT Hunter 2.0" &vbNewLine
WScript.Quit

elseIf intArgCount = 4 And LCase(objArgs.Item(0)) = "/c" Then 

WScript.Echo "Starting comand line mode" &vbNewLine
WScript.Echo "Please make sure to pass the arguments as username->password->target lists file" &vbNewLine

	username = objArgs.Item(1)
	password = objArgs.Item(2)
	inputFile  = objArgs.Item(3)
	If FSO.FileExists (inputFile) Then
	WScript.Echo "Host file exists"
	else
	WScript.Echo "Could not locate the host file ! Try again !"
	WScript.Quit
	END IF
	
else

	WScript.Echo "Syntax error ! Please refer to the help page by typing [cscript APT_Hunter.vbs  -h] or [cscript APT_Hunter.vbs  --help]  "
	WScript.Quit
	
End If


'################################################################################################

On Error GoTo 0

OutPutfolder = strFolder & "\Output"
DataFolder = strFolder & "\Data"
ToolFolder = strFolder & "\Tools"

If FSO.FolderExists(OutPutfolder) Then
	Wscript.Echo "The OutPutfolder exists."
Else
	FSO.CreateFolder(OutPutfolder)
End If

If FSO.FolderExists(DataFolder) Then
	Wscript.Echo "The DataFolder exists."
Else
	FSO.CreateFolder(DataFolder)
End If

If FSO.FolderExists(ToolFolder) Then
	Wscript.Echo "The ToolFolder exists."
Else
	WScript.Echo "Could not locate the tools folder!"
	WScript.Quit	
End If

localadminfolder = DataFolder & "\localadmins"
C_drive_ACLs_folder= DataFolder & "\C_drive_ACLs"
hostfilesfolder = DataFolder & "\hostfiles"
shimcachefolder = DataFolder & "\shimcache"
atjobsfolder = DataFolder & "\atjobs"
AmCachefolder = DataFolder & "\AmCache"
RecentFileCachefolder = DataFolder & "\RecentFileCache"

outputfilePath = OutPutfolder & "\back_doors_check.csv"
outputfile1Path = OutPutfolder & "\APT_preferred_processes_dump.csv"
outputfile2Path = OutPutfolder & "\logs.csv"
outputfile3Path = OutPutfolder & "\services_dump.csv"
outputfile5Path = OutPutfolder & "\network_connections_tcpvcon.csv"
outputfile6Path = OutPutfolder & "\network_connections_netstat.csv"
outputfile7Path = OutPutfolder & "\rogue_processes_check.csv"
outputfile8Path = OutPutfolder & "\shimcache_sweep_2008_2012.tsv"
outputfile9Path = OutPutfolder & "\autorun_startup_command.csv"
outputfile10Path = OutPutfolder & "\localadmin.csv"
outputfile11Path = OutPutfolder & "\localdns.csv"
outputfile12Path = OutPutfolder & "\systeminfo.csv"
outputfile13Path = OutPutfolder & "\RDP_port_check.csv"
outputfile14Path = OutPutfolder & "\systemroot_executables_listing.csv"
outputfile15Path = OutPutfolder & "\prefetch_files_listing.csv"
outputfile16Path = OutPutfolder & "\PsLoggedon.csv"
outputfile17Path = OutPutfolder & "\C_Drive_ACLs.csv"
outputfile18Path = OutPutfolder & "\AtJobs.csv"
outputfile19Path = OutPutfolder & "\shimcache_sweep_2003.tsv"
outputfile20Path = OutPutfolder & "\amcache_sweep.tsv"
outputfile21Path = OutPutfolder & "\RecentFileCache_sweep.csv"
tcpvconSrcPath = ToolFolder & "\tcpvcon.exe"
RawCopySrcPath = ToolFolder & "\Rawcopy\RawCopy64.exe"
shimcacheSrcPath = ToolFolder & "\AppCompatCacheParser\AppCompatCacheParser.exe"
shimcacheSrcPath1 = ToolFolder & "\ShimCacheParser\ShimCacheParser.exe"
jobParsePath =  ToolFolder & "\HCTOOLS\jobparse.exe"
AmCacheSrcPath = ToolFolder & "\AmCacheParser\AmcacheParser.exe"
RfcSrcPath = ToolFolder & "\HCTOOLS\rfc.exe"


If FSO.FileExists (outputfilePath) Then
	'Wscript.Echo "back_doors_check.csv exists."
	Set outputfile = FSO.OpenTextFile(OutPutfolder & "\back_doors_check.csv", ForAppending, True)
Else
	Set outputfile = FSO.OpenTextFile(OutPutfolder & "\back_doors_check.csv", ForAppending, True)
	outputfile.write "HostName,RogueSethcFileCheck1,RogueUtilManFileCheck1,RogueSethcFileCheck2,RogueUtilManFileCheck2,RogueSethcRegistryCheck,RogueUtilManRegistryCheck,RogueOskRegistryCheck,RogueNarratorRegistryCheck,RogueMagnifyRegistryCheck,RogueDisplaySwitchRegistryCheck,AtJobCheck,PsexecCheck, WDigestReg Mimikatz Backdoor"
End If

If FSO.FileExists (outputfile1Path) Then
	'Wscript.Echo "APT_preferred_processes_dump.csv exists."
	Set outputfile1 = FSO.OpenTextFile(OutPutfolder & "\APT_preferred_processes_dump.csv", ForAppending, True)
Else
	Set outputfile1 = FSO.OpenTextFile(OutPutfolder & "\APT_preferred_processes_dump.csv", ForAppending, True)
	outputfile1.write "HostName" & ",ProcessName" & ",ProcessID" & ",ProcessPath" & ",ProcessArguments" & ",ProcessPriority" & ",ProcessOwner" & ",ProcessCreationTime" & ",ParentProcessCreationTime" & ",ParentProcessName" & ",ServiceName" & ",ServiceCaption" & ",ServiceState" & ",ServiceStartMode" & ",ServiceType" & ",ServiceErrorControl" & ",ServicePathName" & VBNewLine
End If

If FSO.FileExists (outputfile2Path) Then
	'Wscript.Echo "logs.csv exists."
	Set outputfile2 = FSO.OpenTextFile(OutPutfolder & "\logs.csv", ForAppending, True)
Else
	Set outputfile2 = FSO.OpenTextFile(OutPutfolder & "\logs.csv", ForAppending, True)
	outputfile2.write "HostName, Start Time"
End If

If FSO.FileExists (outputfile3Path) Then
	'Wscript.Echo "services_dump.csv exists."
	Set outputfile3 = FSO.OpenTextFile(OutPutfolder & "\services_dump.csv", ForAppending, True)
Else
	Set outputfile3 = FSO.OpenTextFile(OutPutfolder & "\services_dump.csv", ForAppending, True)
	outputfile3.write "HostName" & ",ServiceName" & ",ServiceCaption" & ",ServiceState" & ",ServiceStartMode" & ",ServiceType" & ",ServiceErrorControl" & ",ServicePathName" & VBNewLine
End If

If FSO.FileExists (outputfile5Path) Then
	'Wscript.Echo "network_connections_tcpvcon.csv exists."
	Set outputfile5 = FSO.OpenTextFile(OutPutfolder & "\network_connections_tcpvcon.csv", ForAppending, True)
Else
	Set outputfile5 = FSO.OpenTextFile(OutPutfolder & "\network_connections_tcpvcon.csv", ForAppending, True)
	outputfile5.write "HostName,Protocol,ProcessName,ProcessID,Status,SrcIP,DstIP"& VBNewLine
End If


If FSO.FileExists (outputfile7Path) Then
	'Wscript.Echo "rogue_processes_check.csv exists."
	Set outputfile7 = FSO.OpenTextFile(OutPutfolder & "\rogue_processes_check.csv", ForAppending, True)
Else
	Set outputfile7 = FSO.OpenTextFile(OutPutfolder & "\rogue_processes_check.csv", ForAppending, True)
	outputfile7.write "HostName,PID,ProcessName,ProcessPath,ProcessArguments,ProcessOwner,ServiceName,ServiceCaption,ServiceState,ServiceStartMode,ServiceType,ServiceErrorControl,ServicePathName"& VBNewLine
End If

If FSO.FileExists (outputfile8Path) Then
	'Wscript.Echo "shimcache_sweep_2008_2012.tsv exists."
	Set outputfile8 = FSO.OpenTextFile(OutPutfolder & "\shimcache_sweep_2008_2012.tsv", ForAppending, True)
Else
	Set outputfile8 = FSO.OpenTextFile(OutPutfolder & "\shimcache_sweep_2008_2012.tsv", ForAppending, True)
	outputfile8.write "HostName" & vbTab & "ControlSet" & vbTab & "CacheEntryPosition" & vbTab & "File Path" & vbTab & "Last Modified" & vbTab & "Execution"  & VBNewLine
End If

If FSO.FileExists (outputfile9Path) Then
	'Wscript.Echo "autorun_startup_command.csv exists."
	Set outputfile9 = FSO.OpenTextFile(OutPutfolder & "\autorun_startup_command.csv", ForAppending, True)
Else
	Set outputfile9 = FSO.OpenTextFile(OutPutfolder & "\autorun_startup_command.csv", ForAppending, True)
	outputfile9.write "HostName,Caption,Command,Description,Location,Name,SettingID,User,UserSID"& VBNewLine
End If

If FSO.FileExists (outputfile10Path) Then
	'Wscript.Echo "localadmin.csv exists."
	Set outputfile10 = FSO.OpenTextFile(OutPutfolder & "\localadmin.csv ", ForAppending, True)
Else
	Set outputfile10 = FSO.OpenTextFile(OutPutfolder & "\localadmin.csv ", ForAppending, True)
	outputfile10.write "HostName, adminName"& VBNewLine
End If

If FSO.FileExists (outputfile11Path) Then
	'Wscript.Echo "localdns.csv exists."
	Set outputfile11 = FSO.OpenTextFile(OutPutfolder & "\localdns.csv ", ForAppending, True)
Else
	Set outputfile11 = FSO.OpenTextFile(OutPutfolder & "\localdns.csv ", ForAppending, True)
	outputfile11.write "HostName" & vbTab & "LocalDNSinfo"& VBNewLine
End If

If FSO.FileExists (outputfile12Path) Then
	'Wscript.Echo "systeminfo.csv exists."
	Set outputfile12 = FSO.OpenTextFile(OutPutfolder & "\systeminfo.csv", ForAppending, True)
Else
	Set outputfile12 = FSO.OpenTextFile(OutPutfolder & "\systeminfo.csv", ForAppending, True)
	outputfile12.write "HostName,Name,OSType,Description,CurrentTimeZone,OSInstallDate,LastBootTime,MemorySize,Version,Caption" & VBNewLine
End If

If FSO.FileExists (outputfile13Path) Then
	'Wscript.Echo "RDP_port_check.csv exists."
	Set outputfile13 = FSO.OpenTextFile(OutPutfolder & "\RDP_port_check.csv ", ForAppending, True)
Else
	Set outputfile13 = FSO.OpenTextFile(OutPutfolder & "\RDP_port_check.csv ", ForAppending, True)
	outputfile13.write "HostName, RDP_Port"& VBNewLine
End If

If FSO.FileExists (outputfile14Path) Then
	'Wscript.Echo "systemroot_executables_listing.csv exists."
	Set outputfile14 = FSO.OpenTextFile(OutPutfolder & "\systemroot_executables_listing.csv", ForAppending, True)
Else
	Set outputfile14 = FSO.OpenTextFile(OutPutfolder & "\systemroot_executables_listing.csv", ForAppending, True)
	outputfile14.write "HostName, FileName, FileCreationTime, FileModificationTime, FileAccessTime,File Size"& VBNewLine
End If

If FSO.FileExists (outputfile15Path) Then
	'Wscript.Echo "prefetch_files_listing.csv exists."
	Set outputfile15 = FSO.OpenTextFile(OutPutfolder & "\prefetch_files_listing.csv", ForAppending, True)
Else
	Set outputfile15 = FSO.OpenTextFile(OutPutfolder & "\prefetch_files_listing.csv", ForAppending, True)
	outputfile15.write "HostName, FileName, FileCreationTime, FileModificationTime, FileAccessTime,File Size"& VBNewLine
End If


If FSO.FileExists (outputfile17Path) Then
	'Wscript.Echo "C_Drive_ACLs.csv exists."
	Set outputfile17 = FSO.OpenTextFile(OutPutfolder & "\C_Drive_ACLs.csv", ForAppending, True)
Else
	Set outputfile17 = FSO.OpenTextFile(OutPutfolder & "\C_Drive_ACLs.csv", ForAppending, True)
	outputfile17.write "HostName" & ",ACLs" & VBNewLine
End If

If FSO.FileExists (outputfile18Path) Then
	'Wscript.Echo "AtJobs.csv exists."
	Set outputfile18 = FSO.OpenTextFile(OutPutfolder & "\AtJobs.csv", ForAppending, True)
Else
	Set outputfile18 = FSO.OpenTextFile(OutPutfolder & "\AtJobs.csv", ForAppending, True)
	outputfile18.write "HostName, Time, Command, Status"  & VBNewLine
End If

If FSO.FileExists (outputfile19Path) Then
	'Wscript.Echo "shimcache_sweep_2003.tsv exists."
	Set outputfile19 = FSO.OpenTextFile(OutPutfolder & "\shimcache_sweep_2003.tsv", ForAppending, True)
Else
	Set outputfile19 = FSO.OpenTextFile(OutPutfolder & "\shimcache_sweep_2003.tsv", ForAppending, True)
	outputfile19.write "HostName" & vbTab & "Last Modified" & vbTab & "Last Update" & vbTab & "Path" & vbTab & "File Size" & vbTab & "Process Exec Flag" & VBNewLine  
End If

If FSO.FileExists (outputfile20Path) Then
	'Wscript.Echo "amcache_sweep.tsv exists."
	Set outputfile20 = FSO.OpenTextFile(OutPutfolder & "\amcache_sweep.tsv", ForAppending, True)
Else
	Set outputfile20 = FSO.OpenTextFile(OutPutfolder & "\amcache_sweep.tsv", ForAppending, True)
	outputfile20.write "HostName" & vbTab & "ProgramName" & vbTab & "ProgramID" & vbTab & "VolumeID" & vbTab & "VolumeIDLastWriteTimestamp" & vbTab & "PFileID" & vbTab & "FileIDLastWriteTimestamp" & vbTab & "SHA1" & vbTab & "FullPath" & vbTab & "FileExtension" & vbTab & "FileSize" & vbTab & "FileVersionString" &  vbTab & "FileVersionNumber" & vbTab & "FileDescription" & vbTab & "PEHeaderSize" & vbTab & "PEHeaderHash" & vbTab & "PEHeaderChecksum" &  vbTab & "Created" & vbTab & "LastModified" & vbTab & "LastModified2" & vbTab & "CompileTime" & vbTab & "LanguageID" &  VBNewLine  
End If

If FSO.FileExists (outputfile21Path) Then
	'Wscript.Echo "RecentFileCache_sweep.csv exists."
	Set outputfile21 = FSO.OpenTextFile(OutPutfolder & "\RecentFileCache_sweep.csv", ForAppending, True)
Else
	Set outputfile21 = FSO.OpenTextFile(OutPutfolder & "\RecentFileCache_sweep.csv", ForAppending, True)
	outputfile21.write "HostName, FilePath"  & VBNewLine  
End If


If Not FSO.FolderExists(hostfilesfolder) Then

	FSO.CreateFolder(hostfilesfolder)
End If

If Not FSO.FolderExists(localadminfolder) Then

	FSO.CreateFolder(localadminfolder)
End If

If Not FSO.FolderExists(shimcachefolder) Then

	FSO.CreateFolder(shimcachefolder)
End If

If Not FSO.FolderExists(atjobsfolder) Then

	FSO.CreateFolder(atjobsfolder)
End If

If Not FSO.FolderExists(AmCachefolder) Then

	FSO.CreateFolder(AmCachefolder)
End If

If Not FSO.FolderExists(RecentFileCachefolder) Then

	FSO.CreateFolder(RecentFileCachefolder)
End If

If Not FSO.FolderExists(C_drive_ACLs_folder) Then

	FSO.CreateFolder(C_drive_ACLs_folder)
End If

If FSO.FileExists (inputfile) Then
	Set objFileInput = FSO.OpenTextFile(inputfile, 1)
	Do While objFileInput.AtEndOfStream = False
		Server = objFileInput.ReadLine()
		outputfile2.write vbnewline
		outputfile2.write server
		ServerShare = "\\" & Server & "\c$"
		winname = "windows"
		winpath = ServerShare & "\" & winname
		stickykeypath = ServerShare & "\" & winname & "\system32"
		stickykeypath1 = ServerShare & "\" & winname & "\system32\dllcache"
		stickykeypathSethc = ServerShare & "\" & winname & "\system32\sethc.exe"
		stickykeypathUtilman = ServerShare & "\" & winname & "\system32\utilman.exe"
		stickykeypath1Sethc = ServerShare & "\" & winname & "\system32\dllcache\sethc.exe"
		stickykeypath1Utilman = ServerShare & "\" & winname & "\system32\dllcache\utilman.exe"
		atjobpath = ServerShare & "\" & winname & "\tasks"
		prefetchpath = ServerShare & "\" & winname & "\prefetch"
		atjobfilepath = ServerShare & "\" & winname & "\tasks\*.job"
		atjobssubfolder = atjobsfolder & "\" & server
		localAtJobsFilePath = atjobsfolder & "\" & server & "_atjob.csv"
		psexesvcpath = ServerShare & "\" & winname & "\psexesvc.exe"
		hostfilepath = ServerShare & "\" & winname & "\system32" & "\drivers" & "\etc" & "\hosts"
		localhostfilepath = hostfilesfolder & "\"  & Server & "." & "host"
		localPsloggedOnfilepath = PsLoggedonfolder & "\"  & Server & "." & "loggedon"
		tempfolderPath =  ServerShare & "\" & winname & "\temp\" 
		amCacheFilePath = ServerShare & "\" & winname & "\AppCompat\Programs\Amcache.hve"
		amCacheCopyPath =   "C:\" & winname & "\AppCompat\Programs\Amcache.hve"
		RecentFileCacheFilePath = ServerShare & "\" & winname & "\AppCompat\Programs\RecentFileCache.bcf"
		RecentFileCacheCopyPath =   "C:\" & winname & "\AppCompat\Programs\RecentFileCache.bcf"
		tcpvconOutputfile = tempfolderpath & Server & ".netconnection"
		tcpvconDstPath = tempfolderpath & "tcpvcon.exe"
		RawcopyDstPath = tempfolderpath & "RawCopy64.exe"
		localAdminOutputfile = tempfolderpath & Server & ".localadmin"
		LocalAdminLocalfile = localadminfolder & "\"  & Server & "." & "localadmin"
		CDriveACLsOutputfile = tempfolderpath & Server & ".acls"
		CDriveACLsLocalfile = C_drive_ACLs_folder & "\"  & Server & "." & "acls"
		localShimCacheOutputfile = shimcachefolder & "\" & Server
		localShimCacheOutputfile1 = shimcachefolder & "\" & Server &  "_shimcache.tsv"
		netStatOutputfile = tempfolderpath & Server & ".netstat"
		systemhivefile = tempfolderpath & Server & "_system.reg"
		AmCacheFile = tempfolderpath & "Amcache.hve"
		LocalAmCacheOutputFile =  AmCachefolder & "\" & Server
		RecentFileCacheFile = tempfolderpath & "RecentFileCache.bcf"
		LocalRecentFileCacheOutputFile = RecentFileCachefolder & "\" & Server &  "_RecentFileCache.csv"
		
		wscript.echo server & " start at " & FormatDateTime(now)
		outputfile2.write "," & FormatDateTime(now)
		Set NetworkObject = CreateObject("WScript.Network")
		On Error Resume Next
		
		pingtest = pingcheck (server)
	If pingtest = true then
			
			Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
			Set objSWbemService = objSWbemLocator.ConnectServer(Server,"root\cimv2","" & UserName & "",password)			

			If Err.Number <> 0 Then			   
			Wscript.echo Server & " WMI CIMV2 fail"
			outputfile2.write "," & "not able to launch WMI CIMV2"  
			Else
			Wscript.echo Server & " WMI CIMV2 success"
			outputfile2.write "," & "successfully launch WMI CIMV2"  	
				
			NetworkObject.MapNetworkDrive "", ServerShare, False, "" & UserName & "", Password
			
		If Err.Number <> 0 Then
				
				Wscript.echo Server & " SMB fail"
				outputfile2.write "," & "not able to mount SMB share" 
				
			Else
				
				Wscript.echo Server & " SMB success"
				outputfile2.write "," & "successfully mount SMB share" 	
				
				If Not FSO.FolderExists (winpath) Then
			
					winname = "winnt"
					winpath = ServerShare & "\" & winname
		            stickykeypath = ServerShare & "\" & winname & "\system32"
		            stickykeypath1 = ServerShare & "\" & winname & "\system32\dllcache"
		            stickykeypathSethc = ServerShare & "\" & winname & "\system32\sethc.exe"
		            stickykeypathUtilman = ServerShare & "\" & winname & "\system32\utilman.exe"
		            stickykeypath1Sethc = ServerShare & "\" & winname & "\system32\dllcache\sethc.exe"
		            stickykeypath1Utilman = ServerShare & "\" & winname & "\system32\dllcache\utilman.exe"
		            atjobpath = ServerShare & "\" & winname & "\tasks"
					prefetchpath = ServerShare & "\" & winname & "\prefetch"
		            atjobfilepath = ServerShare & "\" & winname & "\tasks\*.job"
		            psexesvcpath = ServerShare & "\" & winname & "\psexesvc.exe"
		            hostfilepath = ServerShare & "\" & winname & "\system32" & "\drivers" & "\etc" & "\hosts"
		            localhostfilepath = hostfilesfolder & "\"  & Server & "." & "host"
		            tempfolderpath =  ServerShare & "\" & winname & "\temp\"
					tcpvconOutputfile = tempfolderpath & Server & ".netconnection"
					localAdminOutputfile = tempfolderpath & Server & ".localadmin"
		            tcpvconDstPath = tempfolderpath & "tcpvcon.exe"
					netStatOutputfile = tempfolderpath & Server & ".netstat"
		            systemhivefile = tempfolderpath & c & "_system.reg"
					CDriveACLsOutputfile = tempfolderpath & Server & ".acls"
					'wscript.echo "working path:" & winname
				End If						
				
							
				Set objSWbemService1 = objSWbemLocator.ConnectServer(Server,"root\default","" & UserName & "",password)
				
				If Err.Number <> 0 Then			   
				Wscript.echo Server & " WMI DEFAULT fail"
				outputfile2.write "," & "not able to launch WMI DEFAULT"  
			    Else
				Wscript.echo Server & " WMI DEFAULT success"
				outputfile2.write "," & "successfully launch WMI DEFAULT"  
				
			    End if
			
			   Set objProcess = objSWbemService.Get("Win32_Process")
'++++++++The modules executed for each IP:++++++++++++
			WMIExec()
			Wscript.echo Server & " system backdoors huntinng started"
			WMIFileCheck(server)
			Wscript.echo Server & " system backdoors huntinng finished"
			Wscript.echo Server & " system information enumeration started"
			WMIQueryLocalAdmins(server)
			WMIQuerySystemInfo()
			WMIQueryProcess ("cmd.exe")
			WMIQueryProcess ("svchost.exe")
			WMIQueryProcess ("explorer.exe")
			WMIQueryProcess ("iexplore.exe")
			WMIQueryProcess ("dllhost.exe")
			WMIQueryProcess ("lsass.exe")
			WMIQueryProcess ("notepad.exe")
			WMIQueryProcess ("rundll32.exe")
			WMIQueryProcess ("winlogon.exe")
			WMIQueryProcess3 ("rundll32.exe")
			WMIQueryProcess3 ("explorer.exe")
			WMIQueryAutoRun()
			WMIQueryService()
			WMIQueryRogueProcess()
			FSOListSystemRootExecutables(winpath)
			FSOReadHostfile()
			FSOReadTcpvconfile()
			FSOReadLocalAdminfile()
			FSOReadCDriveACLs() 
			Wscript.echo Server & " system information enumeration finished"
			Wscript.echo Server & " lateral movement & privilege escalation hunting  started"
			FSOListPrefetchFiles(prefetchpath)
			FSOReadShimcache_method1() 
			FSOReadShimcache_method2()	
			FSOReadAmcache()
			FSOReadRecentFileCache()	
			Wscript.echo Server & " lateral movement & privilege escalation hunting  finished"			
			WMIFSOHostCleanUp ()
			Wscript.echo Server & " local data prcesssing started"
			ParseShimcache1() 
			ParseShimcache2() 
			ParseAmCache() 
			ParseRecentFileCache()
			ParseLocalAdminfile() 
			ParseHostfile()
			ParseCDriveACLs()
			Wscript.echo Server & " local data prcesssing finished"
			Set objSWbemService1 = Nothing
			Set objSWbemService = Nothing
			NetworkObject.RemoveNetworkDrive ServerShare, True, False
			
		End if 
	End If
End if
wscript.echo server & " finish " &  FormatDateTime(now)
outputfile2.write "," & FormatDateTime(now)
On Error GoTo 0

Loop
objFileInput.Close()
End if

'++++++++This function is used to check if the target is live:++++++++++++

Function pingcheck (strComputer)
Dim objShell,objScriptExec
Set objShell = CreateObject("WScript.Shell")
Set objScriptExec = objShell.Exec( _
"ping -n 2 -w 1000 " & strComputer)
strPingResults = LCase(objScriptExec.StdOut.ReadAll)
livehost = true
If InStr(strPingResults, "reply from") Then
If InStr(strPingResults, "destination net unreachable") Then
	WScript.Echo strComputer & "did not respond to ping."
	livehost = false
Else
	WScript.Echo strComputer & " responded to ping."
End If 
Else
WScript.Echo strComputer & " did not respond to ping."
livehost = false
End If
pingcheck = livehost
Set objShell = nothing
Set objScriptExec = nothing
End Function

'++++++++This function is used to execute command on the target remotely++++++++++++

Function WMIExec()

If Not FSO.FileExists (tcpvconSrcPath) Then

Wscript.Echo "tcpvcon.exe does not exist in the current folder."

End If


If FSO.FolderExists(tempfolderpath) Then
'Wscript.Echo "The temp folder  exists"
FSO.CopyFile tcpvconSrcPath,tempfolderpath,true
If Err.Number <> 0 Then
	outputfile2.write ",failed to copy tcpvcon exe"
	wscript.echo Err.Description
else
	outputfile2.write ",successfully copied tcpvcon exe"
	'wscript.echo "successfully copied tcpvcon exe"
End if


Else
FSO.CreateFolder(tempfolderpath)
FSO.CopyFile tcpvconSrcPath,tempfolderpath,true
If Err.Number <> 0 Then
	outputfile2.write ",failed to copy tcpvcon exe"
	wscript.echo Err.Description
else
	outputfile2.write ",successfully copied tcpvcon exe"
	'wscript.echo "successfully copied tcpvcon exe"
End if

End If



strExec = "cmd.exe /c" & tempfolderpath & "tcpvcon.exe -accepteula -anc > " & tempfolderpath & Server & ".netconnection"
intReturn = objProcess.Create(strExec, Null, Null, intProcessID)

If intReturn = 0 Then 
'wscript.echo "successfully executed tcpvcon"
outputfile2.write ",successfully executed tcpvcon"
else 
'wscript.echo "failed to execute tcpvcon"
outputfile2.write ",failed to execute tcpvcon"

End If


strExec = "cmd.exe /c net localgroup administrators > " & tempfolderpath & Server & ".localadmin"
intReturn = objProcess.Create(strExec, Null, Null, intProcessID)

If intReturn = 0 Then 
'wscript.echo "successfully listed local administrators"
outputfile2.write ",successfully listed local administrators"
else 
'wscript.echo "failed to list local administrators"
outputfile2.write ",failed to list local administrators"

End If

strExec = "cmd.exe /c cacls " & winpath&  " > " & tempfolderpath & Server & ".acls"
intReturn = objProcess.Create(strExec, Null, Null, intProcessID)

If intReturn = 0 Then 
'wscript.echo "successfully enumerated C drive ACLs"
outputfile2.write ",successfully enumerated C drive ACLs"
else 
'wscript.echo "failed to enumerateC drive ACLs"
outputfile2.write ",failed to enumerateC drive ACLs"

End If

strExec = "cmd.exe /c reg save HKLM\System " & tempfolderpath & Server & "_system.reg"
intReturn = objProcess.Create(strExec, Null, Null, intProcessID)

If intReturn = 0 Then 
'wscript.echo "successfully export SYSTEM hive"
outputfile2.write ",successfully export SYSTEM hive"
else 
'wscript.echo "failed to export SYSTEM hive"
outputfile2.write ",failed to export SYSTEM hive"

End If

If FSO.FileExists(amCacheFilePath) Then
'Wscript.Echo "The Amcache.hve exists"
FSO.CopyFile RawCopySrcPath,tempfolderpath,true
If Err.Number <> 0 Then
	outputfile2.write ",failed to copy Rawcopy.exe"
	wscript.echo Err.Description
else
	outputfile2.write ",successfully copied Rawcopy.exe"
	'wscript.echo "successfully copied Rawcopy.exe"
End if

strExec = "cmd.exe /c " & tempfolderpath & "RawCopy64.exe " & amCacheCopyPath & " " & tempfolderpath
intReturn = objProcess.Create(strExec, Null, Null, intProcessID)

If intReturn = 0 Then 
'wscript.echo "successfully copied Amcache"
outputfile2.write ",successfully copied Amcache"
else 
'wscript.echo "failed to copy Amcache"
outputfile2.write ",failed to copy Amcache"

End If

else if FSO.FileExists(RecentFileCacheFilePath) Then

'Wscript.Echo "The RecentFileCache.bcf exists"
FSO.CopyFile RawCopySrcPath,tempfolderpath,true
If Err.Number <> 0 Then
	outputfile2.write ",failed to copy Rawcopy.exe"
	wscript.echo Err.Description
else
	outputfile2.write ",successfully copied Rawcopy.exe"
	'wscript.echo "successfully copied Rawcopy.exe"
End if

strExec = "cmd.exe /c " & tempfolderpath & "RawCopy64.exe " & RecentFileCacheCopyPath & " " & tempfolderpath
intReturn = objProcess.Create(strExec, Null, Null, intProcessID)

If intReturn = 0 Then 
'wscript.echo "successfully copied RecentFileCache"
outputfile2.write ",successfully copied RecentFileCache"
else 
'wscript.echo "failed to copy RecentFileCache"
outputfile2.write ",failed to copy RecentFileCache"

End If

End If

End If

End Function


'++++++++This function is used to read local DNS file from target remotely ++++++++++++

Function FSOReadHostfile() 
Dim objTextFile,outputfile4

Set outputfile4 = FSO.OpenTextFile(hostfilesfolder & "\"  & Server & "." & "host", ForAppending, True)
Set objTextFile = FSO.OpenTextFile _
(hostfilepath, ForReading)
strContents = objTextFile.ReadAll
objTextFile.Close

'Wscript.Echo strContents
outputfile4.write strContents
outputfile4.Close
set objTextFile = nothing
Set outputfile4 = nothing
End Function


'++++++++This function is used to read local DNS  output from target remotely ++++++++++++

Function ParseHostFile() 
Dim objTextFile

If FSO.FileExists (localhostfilepath) Then

Set objTextFile = FSO.OpenTextFile _
(localhostfilepath, ForReading)

Do While objTextFile.AtEndOfStream = False

line =  objTextFile.ReadLine()

pondSignCheck = trim(left(line,1))

If (pondSignCheck <> "#") and IsNumeric(pondSignCheck)then 
	
	'wscript.echo line
    outputfile11.write Server & vbTab & line & VBNewLine
End If 	

Loop

objTextFile.Close
set objTextFile = nothing

Else
'Wscript.Echo "localhostfilepath doesn't exist"
exit function
End If	

End Function


'++++++++This function is used to read the ShimCache data from target remotely via AppCompatCacheParser.exe ++++++++++++
Function FSOReadShimcache_method1() 

Set WshShell = WScript.CreateObject ("WScript.Shell")
If Not FSO.FileExists (shimcacheSrcPath) Then

Wscript.Echo "AppCompatCacheParser.exe does not exist in the current folder."
exit function 
End If


If FSO.FileExists (systemhivefile) Then
'Wscript.Echo "SYSTEM hive exists."
strExec = "cmd.exe /c " & shimcacheSrcPath & " -h " & systemhivefile & " -s " & localShimCacheOutputfile 
'Wscript.Echo strExec
intReturn = WshShell.Run(strExec, 0, true) 

If intReturn = 0 Then 
	'wscript.echo "successfully read shimcache via AppCompatCacheParser.exe"
	outputfile2.write ",successfully read shimcache via AppCompatCacheParser.exe"
	
else 
	'wscript.echo "failed to read shimcache via AppCompatCacheParser.exe "
	outputfile2.write ",failed to read shimcache via AppCompatCacheParser.exe"	
End If
	
Else
'Wscript.Echo "SYSTEM doesn't exist"
exit function
End If	
Set WshShell = Nothing
End Function

'++++++++This function is used to read the ShimCache data from target remotely via ShimCacheParser.exe ++++++++++++
Function FSOReadShimcache_method2() 

Set WshShell = WScript.CreateObject ("WScript.Shell")


If FSO.FileExists (systemhivefile) Then
'Wscript.Echo "SYSTEM hive exists."

If Not FSO.FolderExists (localShimCacheOutputfile) Then


If Not FSO.FileExists (shimcacheSrcPath1) Then

	Wscript.Echo "ShimCacheParser.exe does not exist in the current folder."
exit function 
	
End If

strExec = "cmd.exe /c " & shimcacheSrcPath1 & " -f " & systemhivefile & " -d \t -o " & localShimCacheOutputfile1 
intReturn = WshShell.Run(strExec, 0, true) 
If intReturn = 0 Then 
	'wscript.echo "successfully read shimcache via ShimCacheParser.exe"
	outputfile2.write ",successfully read shimcache via ShimCacheParser.exe"
	
else 
	'wscript.echo "failed to read shimcache via ShimCacheParser.exe"
	outputfile2.write ",failed to read shimcache via ShimCacheParser.exe"	
End If

End If
	
Else
'Wscript.Echo "SYSTEM doesn't exist"
exit function
End If	
Set WshShell = Nothing
End Function

'++++++++This function is used to format ShimCache data gathered from target remotely via AppCompatCacheParser.exe ++++++++++++


Function ParseShimcache1()  

Dim  objFolder, objFolderItem,objTextFile

Set objFolder = FSO.GetFolder(localShimCacheOutputfile)
Set objFolderItem = objFolder.Files

For Each oFile in objFolderItem

If UCase(fso.GetExtensionName(oFile))= "TSV"  Then 

'Wscript.Echo oFile.name
'Wscript.Echo oFile.Path

Set objTextFile = FSO.OpenTextFile _
(oFile.Path, ForReading)

strText = objTextFile.ReadAll

'Wscript.Echo strText

objTextFile.Close

arrLines = Split(strText, vbCrLf)

For i = 1 to (Ubound (arrLines) -2)
    'wscript.echo Server & vbTab & arrLines(i)
	outputfile8.write Server & vbTab & arrLines(i) &  VBNewLine
	
Next

set objTextFile = nothing

End If

Next

Set objFolder = nothing
Set objFolderItem = nothing

End Function


'++++++++This function is used to format ShimCache data gathered from target remotely via ShimCacheParser.exe ++++++++++++

Function ParseShimcache2() 

Dim objTextFile
If FSO.FileExists (localShimCacheOutputfile1) Then
'Wscript.Echo localShimCacheOutputfile1

Dim objStream, strData
Set objStream = CreateObject("ADODB.Stream")
objStream.CharSet = "utf-16"
objStream.Open
objStream.LoadFromFile(localShimCacheOutputfile1)
strData = objStream.ReadText()
arrLines = Split(strData, vbCrLf)

For i = 1 to (Ubound (arrLines) -2)
    'wscript.echo Server & "," & arrLines(i)
	outputfile19.write Server & vbTab & arrLines(i) & VBNewLine
	Next
objStream.Close
set objStream = nothing
Else
exit function
End If	

End Function

'++++++++This function is used to read AmCacheParser.exe output ++++++++++++

Function FSOReadAmcache() 

Set WshShell = WScript.CreateObject ("WScript.Shell")
If Not FSO.FileExists (AmCacheSrcPath) Then

Wscript.Echo "AmCacheParser.exe does not exist in the current folder."
exit function 
End If

'Wscript.Echo AmCacheFile

If FSO.FileExists (AmCacheFile) Then
'Wscript.Echo "AmCacheFile exists."
strExec = "cmd.exe /c " & AmCacheSrcPath & " -f " & AmCacheFile & " -s " & LocalAmCacheOutputFile
'Wscript.Echo strExec
intReturn = WshShell.Run(strExec, 0, true) 

If intReturn = 0 Then 
	'wscript.echo "successfully read AmCache via AmCacheParser.exe"
	outputfile2.write ",successfully read AmCache via AmCacheParser.exe"
	
else 
	'wscript.echo "failed to read AmCache via AmCacheParser.exe "
	outputfile2.write ",failed to read AmCache via AmCacheParser.exe"	
End If
	
Else
'Wscript.Echo "AmCacheFile doesn't exist"
exit function
End If	
Set WshShell = Nothing
End Function



'++++++++This function is used to read RFC.exe output ++++++++++++

Function FSOReadRecentFileCache() 

Set WshShell = WScript.CreateObject ("WScript.Shell")
If Not FSO.FileExists (RfcSrcPath) Then

Wscript.Echo "RFC.exe does not exist in the current folder."
exit function 
End If

'Wscript.Echo RecentFileCacheFile

If FSO.FileExists (RecentFileCacheFile) Then
'Wscript.Echo "RecentFileCacheFile  exists."
strExec = "cmd.exe /c " & RfcSrcPath & " " & RecentFileCacheFile & " > " & LocalRecentFileCacheOutputFile 
'Wscript.Echo strExec
intReturn = WshShell.Run(strExec, 0, true) 

If intReturn = 0 Then 
	'wscript.echo "successfully read RecentFileCache via RFC.exe"
	outputfile2.write ",successfully read RecentFileCache via RFC.exe"
	
else 
	'wscript.echo "failed to read RecentFileCache via RFC.exe "
	outputfile2.write ",failed to read RecentFileCache via RFC.exe"	
End If
	
Else
'Wscript.Echo "RecentFileCacheFile doesn't exist"
exit function
End If	
Set WshShell = Nothing
End Function

'++++++++This function is used to format AmCache data gathered from target remotely via AmCacheParser.exe ++++++++++++


Function ParseAmCache()  

Dim  objFolder, objFolderItem,objTextFile

Set objFolder = FSO.GetFolder(LocalAmCacheOutputFile)
Set objFolderItem = objFolder.Files

For Each oFile in objFolderItem

If UCase(fso.GetExtensionName(oFile))= "TSV"  Then 

'Wscript.Echo oFile.name
'Wscript.Echo oFile.Path

Set objTextFile = FSO.OpenTextFile _
(oFile.Path, ForReading)

strText = objTextFile.ReadAll

'Wscript.Echo strText

objTextFile.Close

arrLines = Split(strText, vbCrLf)

For i = 1 to (Ubound (arrLines) -1)
    'wscript.echo Server & vbTab & arrLines(i)
	outputfile20.write Server & vbTab & arrLines(i) &  VBNewLine
	
Next

set objTextFile = nothing

End If

Next

Set objFolder = nothing
Set objFolderItem = nothing

End Function

'++++++++This function is used to format RecentFileCache data gathered from target remotely via RFC.exe ++++++++++++


Function ParseRecentFileCache()  

Dim objTextFile

If FSO.FileExists (LocalRecentFileCacheOutputFile) Then

Set objTextFile = FSO.OpenTextFile _
(LocalRecentFileCacheOutputFile, ForReading)

strText = objTextFile.ReadAll
objTextFile.Close
arrLines = Split(strText, vbCrLf)

For i = 0 to (Ubound (arrLines) -1)
    'wscript.echo Server & vbTab & arrLines(i)
	outputfile21.write Server & "," & arrLines(i) &  VBNewLine
	
Next

set objTextFile = nothing

End If

End Function

'++++++++This function is used to read tcpview.exe output ++++++++++++
Function FSOReadTcpvconfile() 
Dim objTextFile

Set objTextFile = FSO.OpenTextFile _
(tcpvconOutputfile, ForReading)


Do While objTextFile.AtEndOfStream = False

fName =  objTextFile.ReadLine()

'WScript.Echo Server & "," & fName 
outputfile5.write Server & "," & fName & VBNewLine
Loop

objTextFile.Close

set objTextFile = nothing

End Function


'++++++++This function is used to read NETSTAT output ++++++++++++
Function FSOReadParseNetstatfile() 
Dim objTextFile

Set objTextFile = FSO.OpenTextFile _
(netStatOutputfile, ForReading)

Do While objTextFile.AtEndOfStream = False

line =  objTextFile.ReadLine()

If (trim(left(line,5)) = "TCP") oR (trim(left(line,5)) = "UDP") then 
	
	netpid = trim(right(line,6)) 
	parsedLine = ParseText(Line)
	'wscript.echo parsedLine
	serviceInfo = WMIQueryServiceBasedOnProcessId (netpid)
	processInfo = WMIQueryProcess2 (netpid)
	'wscript.echo Server & "," & parsedLine & "," & processInfo & "," & serviceInfo
	outputfile6.write Server & "," & parsedLine & "," & processInfo & "," & serviceInfo & VBNewLine
End If 	
Loop
objTextFile.Close
outputfile6.Close
set objTextFile = nothing



End Function

'++++++++This function is used to parse and format NETSTAT output ++++++++++++
Function ParseText(Line) 
dim i 

For i = 5 to 2 Step -1 
Line = Replace(Line, String(i, " "), " ") 
Next 

Line = Replace(Line, ":", " ") 
Line = Right(Line, Len(Line) - 1) 
Line = Split(Line, " ") 
ParseText = Line(0) & "," & Line(1) & "," & Line(2) & "," & Line (3) & "," & Line (4) & "," & Line (5) & "," & Line (6)


End Function

'++++++++This function is used to list the members of local administrator group for remote system++++++++++++
Function FSOReadLocalAdminfile() 
Dim objTextFile,tempOutputfile

Set tempOutputfile = FSO.OpenTextFile( localadminfolder & "\"  & Server & "." & "localadmin", ForAppending, True)
Set objTextFile = FSO.OpenTextFile _
(localAdminOutputfile, ForReading)
strContents = objTextFile.ReadAll
objTextFile.Close
'Wscript.Echo strContents
tempOutputfile.write strContents
tempOutputfile.Close
set objTextFile = nothing
Set tempOutputfile = nothing
End Function


Function ParseLocalAdminfile() 
Dim objTextFile

If FSO.FileExists (LocalAdminLocalfile) Then

Set objTextFile = FSO.OpenTextFile _
(LocalAdminLocalfile, ForReading)

strText = objTextFile.ReadAll
objTextFile.Close
arrLines = Split(strText, vbCrLf)

For i = 6 to (Ubound (arrLines) -3)
    'wscript.echo Server & "," & arrLines(i)
	outputfile10.write Server & "," & arrLines(i) & VBNewLine
	
Next

set objTextFile = nothing

	
Else
'Wscript.Echo "LocalAdminLocalfile doesn't exist"
exit function
End If	
End Function


'++++++++This function is used to parse ACLs for the C drive of the remote system++++++++++++
Function FSOReadCDriveACLs() 
Dim objTextFile,tempOutputfile


Set tempOutputfile = FSO.OpenTextFile( C_drive_ACLs_folder & "\"  & Server & "." & "acls", ForAppending, True)
Set objTextFile = FSO.OpenTextFile _
(CDriveACLsOutputfile, ForReading)
strContents = objTextFile.ReadAll
objTextFile.Close
'Wscript.Echo strContents
tempOutputfile.write strContents
tempOutputfile.Close
set objTextFile = nothing
Set tempOutputfile = nothing
End Function


Function ParseCDriveACLs() 
Dim objTextFile

If FSO.FileExists (CDriveACLsLocalfile) Then

Set objTextFile = FSO.OpenTextFile _
(CDriveACLsLocalfile, ForReading)

strText = objTextFile.ReadAll
objTextFile.Close
arrLines = Split(strText, vbCrLf)

For i = 0 to (Ubound (arrLines) -2)
    
	arrLines(i) = trim (replace (arrLines(i), VBCR, VBCRLF))
	'Wscript.Echo arrLines(1)
	'Wscript.Echo arrLines(2)
	'Wscript.Echo arrLines(3)
	if arrLines(i) <> ""  Then
	outputfile17.write Server & "," & arrLines(i) & VBNewLine
	end if
Next

set objTextFile = nothing

Else
'Wscript.Echo "CDriveACLsLocalfile doesn't exist"
exit function
End If	
End Function

'++++++++This function is used to clean up all the files and tools left on the remote system ++++++++++++
Function WMIFSOHostCleanUp ()
strExec = "cmd.exe /c taskkill /F /IM tcpvcon.exe /T"
intReturn = objProcess.Create(strExec, Null, Null, intProcessID)
'Wscript.Echo tcpvconOutputfile

If FSO.FileExists (tcpvconOutputfile) Then
'Wscript.Echo "tcpvcon output file exists."
FSO.DeleteFile tcpvconOutputfile, True

If FSO.FileExists (tcpvconOutputfile) Then
	'Wscript.Echo "failed to delete tcpvcon output file"
	outputfile2.write ",failed to delete tcpvcon output file"
Else
	'Wscript.Echo "successfully deleted tcpvcon output file"
	outputfile2.write ",successfully deleted tcpvcon output file"
End If	


End If

If FSO.FileExists (tcpvconDstPath) Then
'Wscript.Echo "tcpvcon.exe exists."
FSO.DeleteFile tcpvconDstPath, True

If FSO.FileExists (tcpvconDstPath) Then
	'Wscript.Echo "failed to delete tcpvcon.exe"
	outputfile2.write ",failed to delete tcpvcon.exe"
Else
	'Wscript.Echo "successfully deleted tcpvcon.exe"
	outputfile2.write ",successfully deleted tcpvcon.exe"
End If	


End If

strExec = "cmd.exe /c taskkill /F /IM rawcopy64.exe /T"
intReturn = objProcess.Create(strExec, Null, Null, intProcessID)
'Wscript.Echo tcpvconOutputfile

If FSO.FileExists (RawcopyDstPath) Then
'Wscript.Echo "Rawcopy64.exe file exists."
FSO.DeleteFile RawcopyDstPath, True

If FSO.FileExists (RawcopyDstPath) Then
	'Wscript.Echo "failed to delete Rawcopy64.exe"
	outputfile2.write ",failed to delete Rawcopy64.exe"
Else
	'Wscript.Echo "successfully deleted Rawcopy64.exe"
	outputfile2.write ",successfully deleted Rawcopy64.exe"
End If	


End If


If FSO.FileExists (localAdminOutputfile) Then
'Wscript.Echo "localAdminOutputfile exists."

FSO.DeleteFile localAdminOutputfile, True
If FSO.FileExists (localAdminOutputfile) Then
	'Wscript.Echo "failed to delete localAdminOutputfile"
	outputfile2.write ",failed to delete localAdminOutputfile"
Else
	'Wscript.Echo "successfully deleted localAdminOutputfile"
	outputfile2.write ",successfully deleted localAdminOutputfile"
End If

End If


If FSO.FileExists (CDriveACLsOutputfile) Then
'Wscript.Echo "CDriveACLsOutputfile exists."

FSO.DeleteFile CDriveACLsOutputfile, True
If FSO.FileExists (CDriveACLsOutputfile) Then
	'Wscript.Echo "failed to delete CDriveACLsOutputfile"
	outputfile2.write ",failed to delete CDriveACLsOutputfile"
Else
	'Wscript.Echo "successfully deleted CDriveACLsOutputfile"
	outputfile2.write ",successfully deleted CDriveACLsOutputfile"
End If

End If

If FSO.FileExists (netStatOutputfile) Then
'Wscript.Echo "netStatOutputfile exists."

FSO.DeleteFile netStatOutputfile, True
If FSO.FileExists (netStatOutputfile) Then
	'Wscript.Echo "failed to delete netStatOutputfile"
	outputfile2.write ",failed to delete netStatOutputfile"
Else
	'Wscript.Echo "successfully deleted netStatOutputfile"
	outputfile2.write ",successfully deleted netStatOutputfile"
End If

End If

If FSO.FileExists (systemhivefile) Then
'Wscript.Echo "systemhivefile exists."

FSO.DeleteFile systemhivefile, True
If FSO.FileExists (systemhivefile) Then
	'Wscript.Echo "failed to delete systemhivefile"
	outputfile2.write ",failed to delete systemhivefile"
Else
	'Wscript.Echo "successfully deleted systemhivefile"
	outputfile2.write ",successfully deleted systemhivefile"
End If

End If

If FSO.FileExists (AmCacheFile) Then
'Wscript.Echo "AmCacheFile exists."

FSO.DeleteFile AmCacheFile, True
If FSO.FileExists (AmCacheFile) Then
	'Wscript.Echo "failed to delete AmCacheFile"
	outputfile2.write ",failed to delete AmCacheFile"
Else
	'Wscript.Echo "successfully deleted AmCacheFile"
	outputfile2.write ",successfully deleted AmCacheFile"
End If

End If

If FSO.FileExists (RecentFileCacheFile) Then
'Wscript.Echo "RecentFileCacheCopyPath exists."

FSO.DeleteFile RecentFileCacheFile, True
If FSO.FileExists (RecentFileCacheFile) Then
	'Wscript.Echo "failed to delete RecentFileCacheFile"
	outputfile2.write ",failed to delete RecentFileCacheFile"
Else
	'Wscript.Echo "successfully deleted RecentFileCacheFile"
	outputfile2.write ",successfully deleted RecentFileCacheFile"
End If

End If

End Function

'++++++++This function is used to obtain file description from  remote system ++++++++++++

Function GetFileDescription (sFilePath, sProgram)
Dim objShell, objFolder, objFolderItem, i 
If FSO.FileExists(sFilePath & "\" & sProgram) Then
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(sFilePath)
Set objFolderItem = objFolder.ParseName(sProgram)
Dim arrHeaders(300)
For i = 0 To 300
	arrHeaders(i) = objFolder.GetDetailsOf(objFolder.Items, i)
'WScript.Echo i &"- " & arrHeaders(i) & ": " & objFolder.GetDetailsOf(objFolderItem, i)
	If lcase(arrHeaders(i))= "file description" Then
		GetFileDescription = objFolder.GetDetailsOf(objFolderItem, i)
		Exit For
	End If
Next
Set objShell = nothing
Set objFolder = nothing
Set objFolderItem = nothing
End If
End Function

'++++++++This function is used to check if any "At" jobs exist on the remote system ++++++++++++
Function GetSchduleTask (sFilePath)

Dim objShell, objFolder, objFolderItem


Set objShell = CreateObject("Shell.Application")
Set objFolder = FSO.GetFolder(sFilePath)
Set objFolderItem = objFolder.Files

For Each oFile in objFolderItem
If ((UCase(left(oFile.Name, 2)) = "AT") And (fso.GetExtensionName(oFile)="job") ) Then 
	GetSchduleTask = oFile.Name
	Exit For
	
End If

Next

Set objShell = nothing
Set objFolder = nothing
Set objFolderItem = nothing

End Function


'++++++++This function is used to extract the AT job informaiton++++++++++++
Function ReadAtJob(sFolderPath) 

Set WshShell = WScript.CreateObject ("WScript.Shell")
If Not FSO.FileExists (jobParsePath) Then

Wscript.Echo "jobparse.exe does not exist in the current folder."
exit function 
End If

strExec = "cmd.exe /c " & jobParsePath & " -d "  & sFolderPath & " -c > " & localAtJobsFilePath
'wscript.echo strExec
intReturn = WshShell.Run(strExec, 0, true) 
If intReturn = 0 Then 
	'wscript.echo "successfully read AT job files"
	outputfile2.write ",successfully read AT job files"
	
else 
	'wscript.echo "failed to read AT job files"
	outputfile2.write ",failed to read AT job files"	
End If
	

Set WshShell = Nothing
End Function

'++++++++This function is used to parse extracted AT job information ++++++++++++


Function ParseAtJob(sFilePath)   

Dim objTextFile
If FSO.FileExists (sFilePath) Then
'Wscript.Echo sFilePath

Set objTextFile = FSO.OpenTextFile _
(sFilePath, ForReading)

strText = objTextFile.ReadAll

'Wscript.Echo strText

objTextFile.Close

arrLines = Split(strText, vbCrLf)


For i = 0 to (Ubound (arrLines) -1)
    'wscript.echo Server & "," & arrLines(i)
	outputfile18.write Server & "," & arrLines(i) &  VBNewLine
	
Next

set objTextFile = nothing

	
Else
Wscript.Echo sFilePath & " doesn't exist"
exit function
End If

End Function

'+++++++++++++++This function is used to list executables under system root +++++++++++++

Function FSOListSystemRootExecutables (sFilePath)

Dim  objFolder, objFolderItem

Set objFolder = FSO.GetFolder(sFilePath)
Set objFolderItem = objFolder.Files

For Each oFile in objFolderItem
If UCase(fso.GetExtensionName(oFile))= "EXE" or UCase(fso.GetExtensionName(oFile))= "BAT" or UCase(fso.GetExtensionName(oFile))= "VBS" or UCase(fso.GetExtensionName(oFile))= "DLL"  or UCase(fso.GetExtensionName(oFile))= "RAR" or UCase(fso.GetExtensionName(oFile))= "TXT"  Then 

	'Wscript.Echo Server & "," & oFile.Name & "," & oFile.DateCreated  & "," & oFile.DateLastModified & "," & oFile.DateLastAccessed & "," & oFile.Size
	outputfile14.write Server & "," & oFile.Name & "," & oFile.DateCreated  & "," & oFile.DateLastModified & "," & oFile.DateLastAccessed & "," & oFile.Size & VBNewLine
End If

Next
Set objFolder = nothing
Set objFolderItem = nothing

End Function


'+++++++++++++++This function is used to list Prefetch files +++++++++++++

Function FSOListPrefetchFiles (sFilePath)

If FSO.FolderExists (prefetchpath) Then
'wscript.echo prefetchpath  & " exists"

Dim  objFolder, objFolderItem

Set objFolder = FSO.GetFolder(sFilePath)
Set objFolderItem = objFolder.Files

For Each oFile in objFolderItem
If UCase(fso.GetExtensionName(oFile))= "PF"  Then 

	'Wscript.Echo Server & "," & oFile.Name & "," & oFile.DateCreated  & "," & oFile.DateLastModified & "," & oFile.DateLastAccessed & "," & oFile.Size
	outputfile15.write Server & "," & oFile.Name & "," & oFile.DateCreated  & "," & oFile.DateLastModified & "," & oFile.DateLastAccessed & "," & oFile.Size & VBNewLine
End If
Next
Set objFolder = nothing
Set objFolderItem = nothing


End If

End Function
'++++++++This function is used to query detailed service information based on ProcessID++++++++++++

Function WMIQueryServiceBasedOnProcessId(ProcessId)


Dim colItems,objItem

Set colItems = objSWbemService.ExecQuery( "SELECT * FROM Win32_Service"  & " Where ProcessID = '" & ProcessId & "'",,48) 

For Each objItem in colItems 
	
'Wscript.Echo  objItem.Name & ", " & objItem.State & "," & objItem.StartMode & "," & objItem.ServiceType & "," & objItem.ErrorControl & "," &  objItem.PathName & "," & objItem.Caption
WMIQueryServiceBasedOnProcessId = LCase(objItem.Name) & "," & LCase(objItem.Caption) & ", " & LCase(objItem.State) & "," & LCase(objItem.StartMode) & "," & LCase(objItem.ServiceType) & "," & LCase(objItem.ErrorControl) & "," & LCase(objItem.PathName) 
Next

set colItems = nothing 
set objItem = nothing 


End Function




Function WMIDateStringToDate(dtmStart)
WMIDateStringToDate = CDate(Mid(dtmStart, 5, 2) & "/" & _
Mid(dtmStart, 7, 2) & "/" & Left(dtmStart, 4) _
& " " & Mid (dtmStart, 9, 2) & ":" & _
Mid(dtmStart, 11, 2) & ":" & Mid(dtmStart, _
13, 2))
End Function

'++++++++This function is used to query detailed process information from remote system+++++++++++

Function WMIQueryProcess(strProcess)
Dim objProcess, objProcess1, colProcesses, colProcesses1

Set colProcesses = objSWbemService.ExecQuery("Select * from Win32_Process " _
& " Where name = '" & strProcess & "'",,48)


For Each objProcess in colProcesses

colProperties = objProcess.GetOwner( _
strNameOfUser,strUserDomain)
dtmStartTime = objProcess.CreationDate
dtmreturn = WMIDateStringToDate (dtmStartTime)
ParentProcessId = objProcess.ParentProcessId

Set colProcesses1 = objSWbemService.ExecQuery("Select * from Win32_Process " _
& " Where ProcessID = '" & ParentProcessId & "'",,48)


For Each objProcess1 in colProcesses1  

	dtmStartTime1 = objProcess1.CreationDate
	dtmreturn1 = WMIDateStringToDate (dtmStartTime1)
	processService = WMIQueryServiceBasedOnProcessId (objProcess.ProcessId )
	'WScript.StdOut.Write strProcess & "|" & objProcess.ProcessId & "|" & objProcess.ExecutablePath & " | " & objProcess1.name & "|" & objProcess.commandline & "|" &strNameOfUser & "|" & dtmreturn & "|" & dtmreturn1 & "|" & processService & VBNewLine
	outputfile1.Write Server & "," & LCase(strProcess) & "," & LCase(objProcess.ProcessId) & "," & LCase(objProcess.ExecutablePath) & " , " & LCase(objProcess.commandline) & "," & LCase(objProcess.Priority) & "," & LCase(strUserDomain) & "\" & LCase(strNameOfUser) & "," & LCase(dtmreturn) & "," & LCase(dtmreturn1) & "," & LCase(objProcess1.name) & "," & processService & VBNewLine
      
Next
set objProcess1 = nothing 
set colProcesses1 = nothing 


Next
set objProcess = nothing 
set colProcesses = nothing 
End Function

'--------------------------------------------------------------------------------------------------------------


Function WMIQueryProcess2(ProcessID)
Dim objProcess, colProcesses


Set colProcesses = objSWbemService.ExecQuery("Select * from Win32_Process " _
& " Where ProcessID = '" & ProcessID & "'",,48)


For Each objProcess in colProcesses

colProperties = objProcess.GetOwner( strNameOfUser,strUserDomain)
WMIQueryProcess2 =  LCase(objProcess.Name) & "," & LCase(objProcess.ExecutablePath) & "," & LCase(objProcess.commandline) & "," & LCase(strUserDomain) & "/" & LCase(strNameOfUser)

Next
set objProcess = nothing 
set colProcesses = nothing 
End Function

Function WMIQueryProcess3(strProcess)
Dim objProcess, colProcesses

'WScript.echo strProcess

Set colProcesses = objSWbemService.ExecQuery("Select * from Win32_Process " _
& " Where name = '" & strProcess & "'",,48)


For Each objProcess in colProcesses
'WScript.echo objProcess.ProcessId
'WScript.echo objProcess.ExecutablePath
colProperties = objProcess.GetOwner( _
strNameOfUser,strUserDomain)
dtmStartTime = objProcess.CreationDate
dtmreturn = WMIDateStringToDate (dtmStartTime)
processService = WMIQueryServiceBasedOnProcessId (objProcess.ProcessId )

'WScript.StdOut.Write strProcess & "|" & objProcess.ProcessId & "|" & objProcess.ExecutablePath & "|" &strNameOfUser  & VBNewLine
outputfile1.Write Server & "," & LCase(strProcess) & "," & LCase(objProcess.ProcessId) & "," & LCase(objProcess.ExecutablePath) & " , " & LCase(objProcess.commandline) & "," & LCase(objProcess.Priority) & "," & LCase(strUserDomain) & "\" & LCase(strNameOfUser) & "," & LCase(dtmreturn) & "," &  "," &  "," & processService & VBNewLine

Next
set objProcess = nothing 
set colProcesses = nothing 
End Function

'++++++++This function is used to query programs that are not running from common Windows directories +++++++++++
Function WMIQueryRogueProcess()
Dim objProcess, colProcesses


Set colProcesses = objSWbemService.ExecQuery("Select * from Win32_Process Where (NOT ExecutablePath LIKE  '%system32%' And NOT ExecutablePath LIKE  '%syswow64%')  And NOT ExecutablePath LIKE '%Program Files%' ",,48)


For Each objProcess in colProcesses

'colProperties = objProcess.GetOwner( strNameOfUser,strUserDomain)
'WMIQueryProcess2 =  objProcess.Name & "," & objProcess.ExecutablePath & "," & objProcess.commandline & "," & strUserDomain & "/" & strNameOfUser
ProcessId = objProcess.ProcessId
colProperties = objProcess.GetOwner( strNameOfUser,strUserDomain)
serviceinfo = WMIQueryServiceBasedOnProcessId(ProcessId)
'wscript. echo WMIQueryServiceBasedOnProcessId("796")
'Wscript.echo server & "," & objProcess.ProcessId & "," & objProcess.Name & "," & objProcess.ExecutablePath & "," & objProcess.commandline & "," & strUserDomain & "/" & strNameOfUser & "," & serviceinfo
outputfile7.write server & "," & LCase(objProcess.ProcessId) & "," & LCase(objProcess.Name) & "," & LCase(objProcess.ExecutablePath) & "," & LCase(objProcess.commandline) & "," & LCase(strUserDomain) & "/" & LCase(strNameOfUser) & "," & serviceinfo & VBNewLine
Next
set objProcess = nothing 
set colProcesses = nothing 
End Function

'++++++++This function is used to query partial autorun information for remote system  +++++++++++
Function WMIQueryAutoRun()
Dim objAutoRuns, colAutoRuns

Set colAutoRuns = objSWbemService.ExecQuery("Select * from Win32_StartupCommand",,48)

For Each objAutoRuns in colAutoRuns

'Wscript.echo server & "," & objAutoRuns.Caption & "," & objAutoRuns.Command & "," & objAutoRuns.Description & "," & objAutoRuns.Location & "," & objAutoRuns.Name & "," & objAutoRuns.SettingID & "," & objAutoRuns.User & "," & objAutoRuns.UserSID
outputfile9.write server & "," & LCase(objAutoRuns.Caption) & "," & LCase(objAutoRuns.Command) & "," & LCase(objAutoRuns.Description) & "," & LCase(objAutoRuns.Location) & "," & LCase(objAutoRuns.Name) & "," & LCase(objAutoRuns.SettingID) & "," & LCase(objAutoRuns.User) & "," & LCase(objAutoRuns.UserSID) & VBNewLine

Next
set objAutoRuns = nothing 
set colAutoRuns = nothing 
End Function

'++++++++This function is used to check detailed service information on the remote host +++++++++++

Function WMIQueryService()

Dim colItems,objItem

Set colItems = objSWbemService.ExecQuery( _
"SELECT * FROM Win32_Service",,48) 
For Each objItem in colItems 

serviceStartMode = objItem.StartMode
serviceType = objItem.ServiceType
serviceErrorControl = objItem.ErrorControl


'Wscript.Echo  objItem.Name & "| " & objItem.State & "|" & objItem.StartMode & "|" & objItem.ServiceType & "|" & objItem.ErrorControl & "|" &  objItem.PathName & "|" & objItem.Caption
outputfile3.Write Server & "," & LCase(objItem.Name) & ", " & LCase(objItem.Caption) & "," & LCase(objItem.State) & "," & LCase(objItem.StartMode) & "," & LCase(objItem.ServiceType) & "," & LCase(objItem.ErrorControl) & "," &  LCase(objItem.PathName)  & VBNewLine

Next

set colItems = nothing 
set objItem = nothing 


End Function



'++++++++This function is used to query operating system information from remote system  +++++++++++

Function WMIQuerySystemInfo()

Dim colItems,objItem

Set colItems = objSWbemService.ExecQuery( _
"Select * from Win32_OperatingSystem",,48) 
For Each objItem in colItems 


OSInstallDate = WMIDateStringToDate(LCase(objItem.InstallDate))
LastBootTime = WMIDateStringToDate (LCase(objItem.LastBootUpTime))


'Wscript.Echo  Server & "," & LCase(objItem.Name) & "," & LCase(objItem.Caption) & "," & LCase(objItem.OSType) & "," & LCase(objItem.Description) & "," & LCase(objItem.Version) & "," & OSInstallDate & "," & LastBootTime & "," & LCase(objItem.Manufacturer) & "," & LCase(objItem.Organization) & "," & LCase(objItem.TotalVirtualMemorySize) & "," & LCase(objItem.WindowsDirectory) & "," & LCase(objItem.BootDevice) & VBNewLine
outputfile12.Write Server & "," & LCase(objItem.Name) &  "," & LCase(objItem.OSType) & "," & LCase(objItem.Description) & "," & LCase(objItem.CurrentTimeZone) & "," & OSInstallDate & "," & LastBootTime & "," & LCase(objItem.TotalVirtualMemorySize) & "," & LCase(objItem.Version)& "," & LCase(objItem.Caption)  & VBNewLine

Next

set colItems = nothing 
set objItem = nothing 


End Function


'++++++++This function is used to check file properties on the remote system, including StickyKey,Psexec,Atjob, etc +++++++++++


Function WMIFileCheck(ServerName)


outputfile.write vbnewline
outputfile.write servername
'ServerShare = "\\" & ServerName & "\c$"
'winname = "windows"
'winpath = ServerShare & "\" & winname
regsetchpath = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\sethc.exe\debugger"
regutilmanpath = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\utilman.exe\debugger"
strKeyPath1 = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\sethc.exe"
strValueName = "debugger"
strKeyPath2 = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\utilman.exe"
strKeyPath3 = "SYSTEM\CurrentControlSet\Control\Terminal"
strKeyPath4 = "Server\WinStations\RDP-Tcp"
strKeyPath5 = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\osk.exe"
strKeyPath6 = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\Narrator.exe"
strKeyPath7 = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\Magnify.exe"
strKeyPath8 = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\DisplaySwitch.exe"
strKeyPathRDP = strKeyPath3  & " " & strKeyPath4
strValueNameRDP = "PortNumber"
WDigestReg = "SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest"
WDigestRegValue = "UseLogonCredential"
On Error Resume Next

If FSO.FileExists (stickykeypathSethc) Then

strValue3 = GetFileDescription(stickykeypath, "sethc.exe")
If strValue3 <> "Accessibility shortcut keys" And strValue3 <> "Windows NT High Contrast Invocation" Then
outputfile.write "," &  "found:" & strValue3
'wscript.echo strValue3
Else
outputfile.write "," & "stickykey not found"
'wscript.echo strValue3
End If

Else
outputfile.write "," & "stickykey not found"
'wscript.echo "stickykey not found"
End If

If FSO.FileExists (stickykeypathUtilman) Then
strValue4 = GetFileDescription(stickykeypath, "utilman.exe")
If strValue4 <> "UtilMan EXE" And strValue4 <> "Utility Manager"  Then
outputfile.write "," &  "found:" & strValue4
'wscript.echo strValue4
Else
outputfile.write "," & "stickykey not found"
'wscript.echo strValue4
End If
Else
outputfile.write "," & "stickykey not found"
'wscript.echo "stickykey not found"
End If

If FSO.FileExists (stickykeypath1Sethc) Then
strValue5 = GetFileDescription(stickykeypath1, "sethc.exe")
If strValue5 <> "Accessibility shortcut keys" And strValue5 <> "Windows NT High Contrast Invocation" Then
outputfile.write "," &  "found:" & strValue5
'wscript.echo strValue5
Else
outputfile.write "," & "stickykey not found"
'wscript.echo strValue5
End If
Else
outputfile.write "," & "stickykey not found"
'wscript.echo "stickykey not found"
End If

If FSO.FileExists (stickykeypath1Utilman) Then
strValue6 = GetFileDescription(stickykeypath1, "utilman.exe")
If strValue6 <> "UtilMan EXE" And strValue6 <> "Utility Manager"  Then
outputfile.write "," &  "found:" & strValue6
'wscript.echo strValue6
Else
outputfile.write "," & "stickykey not found"
'wscript.echo strValue6
End If
Else
outputfile.write "," & "stickykey not found"
'wscript.echo "stickykey not found"
End If


Set objReg = objSWbemService1.Get("StdRegProv")

objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath1, strValueName, strValue1
objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath2, strValueName, strValue2
objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath5, strValueName, strValue5
objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath6, strValueName, strValue6
objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath7, strValueName, strValue7
objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath8, strValueName, strValue8

objReg.GetDWORDValue HKEY_LOCAL_MACHINE, strKeyPathRDP, strValueNameRDP, strValueRDP
'wscript.echo "-------------------------------------------------------------------------------------------------"

'wscript.echo strValueRDP

'wscript.echo "-------------------------------------------------------------------------------------------------"


objReg.GetDWORDValue HKEY_LOCAL_MACHINE, WDigestReg, WDigestRegValue, strValueWDigest
'wscript.echo "-------------------------------------------------------------------------------------------------"

'wscript.echo strValueWDigest

'wscript.echo "-------------------------------------------------------------------------------------------------"

If IsNull(strValue1) Then
outputfile.write "," & "stickykey not found"
Else
outputfile.write "," & "found:" & chr(34) & strValue1 & chr(34)
'wscript.echo strValue1
End If

If IsNull(strValue5) Then
outputfile.write "," & "stickykey not found"
Else
outputfile.write "," & "found:" & chr(34) & strValue5 & chr(34)
'wscript.echo strValue5
End If

If IsNull(strValue6) Then
outputfile.write "," & "stickykey not found"
Else
outputfile.write "," & "found:" & chr(34) & strValue6 & chr(34)
'wscript.echo strValue6
End If

If IsNull(strValue7) Then
outputfile.write "," & "stickykey not found"
Else
outputfile.write "," & "found:" & chr(34) & strValue7 & chr(34)
'wscript.echo strValue7
End If

If IsNull(strValue8) Then
outputfile.write "," & "stickykey not found"
Else
outputfile.write "," & "found:" & chr(34) & strValue8 & chr(34)
'wscript.echo strValue8
End If

If IsNull(strValue2) Then
outputfile.write "," & "stickykey not found"
Else
outputfile.write "," & "found:" & chr(34) & strValue2 & chr(34)
'wscript.echo strValue2
End If

If IsNull(strValueRDP) Then
'wscript.echo "RDP port not found"
outputfile2.write "," & "RDP port not found"
Else
outputfile2.write "," & "Successfully list RDP port"
outputfile13.write server & "," & strValueRDP & VBNewLine

End If



strValue7 = GetSchduleTask (atjobpath)
If strValue7 <> ""  Then
outputfile.write "," & "found:" & chr(34) & strValue7 & chr(34)
'wscript.echo strValue7



If Not FSO.FolderExists(atjobssubfolder) Then

	FSO.CreateFolder(atjobssubfolder)
End If

FSO.CopyFile atjobfilepath,atjobssubfolder,true

ReadAtJob(atjobssubfolder) 
ParseAtJob(localAtJobsFilePath)

Else
outputfile.write "," & "at job not found"
End If

If FSO.FileExists (psexesvcpath) Then
outputfile.write "," & "found:psexec" 
'wscript.echo psexesvcpath
Else
outputfile.write "," & "psexec not found"
End If


If IsNull(strValueWDigest) Then
'wscript.echo "WDigestReg not found"
outputfile.write "," & "WDigestReg not found"
Else
outputfile.write "," & "WDigestReg found:" & chr(34) & strValueWDigest & chr(34)

End If


'outputfile.Close

On Error GoTo 0

'Set objSWbemService1 = Nothing
'Set objSWbemService = Nothing

End Function

