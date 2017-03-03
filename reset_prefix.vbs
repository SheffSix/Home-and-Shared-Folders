logfile = "reset_prefix_log.txt"
tolog = " 2>>" & logfile
ifLimit = 1
Set ErrorList = CreateObject("System.Collections.ArrayList")

sub speak(smessage)
	if smessage = "" then
		wscript.echo ""
		objShell.Run "%comspec% /C echo. >> " & logfile, 7, true
	else
		wscript.echo smessage
		objShell.Run "%comspec% /C echo " & smessage & " >> " & logfile, 7, true
	end if
end sub

function fResetFolder(sFolder, sUser)
	Return = 1
	ifloop = 0
	Do
		ifLoop = ifLoop + 1
		speak sUser & ": Force ownership to Administrators group"
		Return = objShell.Run ("%comspec% /c takeown /f """ & sFolder & """ /r /d y /a" & tolog, 1, true)
		iReturn = Return
		speak sUser & ": Reset permissions to default"
		Return = objShell.Run ("%comspec% /c icacls """ & sFolder & """ /reset /t" & tolog, 1, true)
		iReturn = iReturn + Return
		speak sUser & ": Grant " & sUser & " access to folder"
		Return = objShell.Run ("%comspec% /c icacls """ & sFolder & """ /grant " & sUser & ":(OI)(CI)F" & tolog, 1, true)
		iReturn = iReturn + Return
		speak sUser & ": Give " & sUser & " ownership"
		Return = objShell.Run ("%comspec% /c icacls """ & sFolder & """ /setowner " & sUser & " /T /C" & tolog, 1, true)
		iReturn = iReturn + Return
		if iReturn = 0 then 
			ifLoop = ifLimit
		else
			speak "Forcing ownership failed. Retrying..."
		End If
	Loop until ifLoop = ifLimit
	if iReturn <> 0 then
		ErrorList.Add sFolder
		speak "ERROR: Forcing ownership did not succeed after " & ifLimit & " attempts."
		speak "       Reset incomplete"
	end if
end function

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = WScript.CreateObject("WScript.Shell")

If wscript.arguments.count = 0 then
	speak "Error - Missing parameters"
	speak "Usage: cscript reset_prefix.vbs \\Path\To\Home\Folder\Root [UserName] [-setroot]"
	speak ""
	wscript.quit
else
	booSetRoot = false
	strFolder = wscript.arguments(0)
	if wscript.arguments.count >= 2 then strSubFolder = wscript.arguments(1)
	if wscript.arguments.count = 3 then
		if wscript.arguments(2) = "-setroot" then booSetRoot = true
	end if
end if

if objFSO.folderExists(strFolder) then
	if right(strFolder, 1) <> "\" and booSetRoot = false then strFolder = strFolder & "\"
	objShell.Run "%comspec% /C echo Reset permissions > " & logfile, 1, true
	if not isempty(strSubFolder) and booSetRoot = false then
		if objFSO.FolderExists(strFolder & strSubFolder) then
			speak "Process Single Subfolder..."
			speak "Processing " & strSubFolder
'wscript.quit
			fresetfolder strFolder & strSubFolder, strSubFolder
		else
			speak "Skipping " & strFolder & strSubFolder & " because it does not exist"
			wscript.quit
		end if
	elseif booSetRoot = true then
		speak "Processing root folder"
'wscript.quit
		fresetfolder strFolder, strSubFolder
	else
		speak "Processing all subfolders"
'wscript.quit
		iloop = 1
		Set objFolder = objFSO.GetFolder(strFolder)
		Set colSubfolders = objFolder.Subfolders
'For Each objSubfolder in colSubfolders
'speak objSubfolder.Name
'next
'wscript.quit
		For Each objSubfolder in colSubfolders
		if iloop > 0 then 'and iloop < 16 then
			strSubFolder = objSubfolder.Name
			speak ""
			Speak "Loopref: " & iloop
			If strSubFolder <> "__DFSR_DIAGNOSTICS_TEST_FOLDER__" then
				speak "Processing " & strSubFolder
				fresetfolder strFolder & strSubFolder, strSubFolder
			else
				speak "Skipping " & strSubFolder & " because it is excluded"
			end if
		end if
		iloop = iloop + 1
		Next
		Set objFolder = Nothing
		Set colSubfolders = Nothing
	End If
	If ErrorList.Count > 0 then
		speak ""
		speak "The following folders failed:"
		for each x in ErrorList
			speak x
		Next
	end if
else
	speak "The folder " & strFolder & " does not exist!"
	wscript.quit
end if
Set objFSO = Nothing
Set objShell = Nothing