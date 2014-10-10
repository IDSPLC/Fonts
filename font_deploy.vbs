' Install multiple Fonts in Windows XP/Vista/7
'This will deploy custom company fonts. To properly track changes, this script needs write access to the following Registry Key: HKLM\Software\IDSPLC\Fonts\
'This registry key tracks the version installed against the version contained in the script. Increasing the Installed version here will cause the script to push out any new fonts, but skip the process if they are the same. This saves time as the script will not have to compare each font to those already installed.


Dim Installed, InstallCheck, RegKey, RegType, RegRun, temp
DIM strFontSourcePath, strComputerName

Installed = 2
RegRun = 0

RegKey = "HKLM\Software\IDSPLC\Fonts\Installed"
RegType = "REG_DWORD"

'strFontSourcePath = "\\US-SVR-NET-001\IDS Inc Shared\IT\Fonts\deploy\"				'Location of fonts to install 

'Debugging Options
on error resume next		'allows script to continue in the event of an error
DisplayOutput = TRUE		'set to True to display script output to command window 
ToOverride = TRUE			'set to True to force overwrite entries even if new value matches those already set

'Header
if DisplayOutput = True then
	Wscript.echo "--------------------------------------"
	Wscript.echo " IDS Custom Font Install"
	Wscript.echo "--------------------------------------"
	Wscript.echo ""
end if


'This section checks to see if the current deployment has already been run
'----------------------------------------------------------
if DisplayOutput = True then
	Wscript.echo "Checking for new fonts"
	Wscript.echo "Official Font Package Level: " & Installed
end if

InstallCheck = ReadReg(RegKey)
if InstallCheck = "" then InstallCheck = 0

if DisplayOutput = True then
	wscript.echo "Current Font Package Level: " & InstallCheck
	wscript.echo ""
end if

If InstallCheck < Installed or ToOverride = TRUE then
	if DisplayOutput = TRUE then
		If ToOverride = TRUE then
			wscript.echo "Override Enabled, Forcing Update"
			wscript.echo ""
		else
			Wscript.echo "Updates Needed"
			wscript.echo ""
		end if
		Wscript.echo "Updating Fonts"
		Wscript.echo ""
	end if
	strFontSourcePath = SetFilePath()
	CopyFonts(strFontSourcePath)
	temp = WriteReg(RegKey,Installed,RegType)
	InstallCheck=ReadReg(RegKey)
	wscript.echo "Updated Font Package Level: " & InstallCheck
	wscript.echo "Update Complete"
Else
	if DisplayOutput = True then
		Wscript.Echo "No Updates Needed"
		Wscript.Echo ""
		WScript.Echo "Exiting"
	end if
End If

Function WriteReg(RegPath, Value, RegType)	'write registry key
	Dim objRegistry, Key
	Set objRegistry = CreateObject("Wscript.shell")
	Key = objRegistry.RegWrite(RegPath, Value, RegType)
	if DisplayOutput = True then
		wscript.echo "Updating Registry"
		wscript.echo ""
	end if
	WriteReg = Key
		if err.number = -2147024894 then 'catch error if registry key does not exist
			if DisplayOutput = True then
			wscript.echo "Registry Key Does Not Exist"
			wscript.echo "Creating New Registry Key " & RegPath
			wscript.echo ""
		end if
		err.clear
		elseif err.number <> 0 then 'catch error if unable to write to registry
			if DisplayOutput = True then
				wscript.echo "Error number: " & err.number & " Source: " & err.source & " Desc: " & err.description
				wscript.echo ""
				wscript.echo "There was an error writing to the registry"
				wscript.echo "Check that you have permission to write to the registry key"
				wscript.echo "Please contact your IT Administrator for assistance"
			end if
		end if
End Function

Function ReadReg(RegPath)	'read registry key
	Dim objRegistry, Key
	Set objRegistry = CreateObject("Wscript.shell")
	Key = objRegistry.RegRead(RegPath)
	if err.number = -2147024894 then
		if DisplayOutput = True then
			wscript.echo "Registry value does not exist"
			wscript.echo ""
		end if
		err.clear
		ReadReg = 0
	elseif err.number <> 0 then 'catch error if unable to read to registry
		if DisplayOutput = True then
			wscript.echo "Error number: " & err.number & " Source: " & err.source & " Desc: " & err.description
			wscript.echo ""
			wscript.echo "There was an error reading from the registry"
			wscript.echo "Please contact your IT Administrator for assistance"
		end if
	else
		ReadReg = Key
	end if	  
End Function

Function CopyFonts(strFontSourcePath)	'Copy fonts to local computer
	Dim objShell, objFSO, wshShell
	Dim objFolder, objFont, objNameSpace, objFile
	Dim counter
	
	Set objShell = CreateObject("Shell.Application")
	Set wshShell = CreateObject("WScript.Shell")
	Set objFSO = createobject("Scripting.Filesystemobject")
	
	counter = 0
	
	if DisplayOutput = True then
		Wscript.Echo "--------------------------------------"
		Wscript.Echo " Install Fonts "
		Wscript.Echo "--------------------------------------"
		Wscript.Echo ""
	end if
	
	If objFSO.FolderExists(strFontSourcePath) Then												'verify that fonts location exists
		If objFSO.FolderExists("C:\Users\") Then												'look for folder not available in XP
			' Start Windows 7/Vista
			if DisplayOutput = True then
				Wscript.Echo "Windows Vista/7 Detected"
				Wscript.Echo ""
			end if
			Set objNameSpace = objShell.Namespace(strFontSourcePath)
			Set objFolder = objFSO.getFolder(strFontSourcePath)
				For Each objFile In objFolder.files
					If LCase(right(objFile,4)) = ".ttf" OR LCase(right(objFile,4)) = ".otf" Then
						Set objFont = objNameSpace.ParseName(objFile.Name)
						If DisplayOutput = TRUE then wscript.echo "Checking Font: " & objFile.Name
						If objFSO.FileExists("C:\WINDOWS\Fonts\" & objFile.Name) = False Then	'Check to see if font is already installed
							if DisplayOutput = TRUE then
								wscript.echo objFile.Name & " not currently installed."
								wscript.echo "Installing " & objFile.Name
								wscript.echo ""
							end if
							objFont.InvokeVerb("Install")										'run font installation routine
							counter = counter + 1
							Set objFont = Nothing
						Else
							If DisplayOutput = True then
								wscript.echo objFile.Name & " already installed"
								wscript.echo ""
							end if
						End If
					End If
				Next
		Else
			' Start Windows XP
			if DisplayOutput = True then
				Wscript.Echo "Windows XP Detected"
				Wscript.Echo ""
			end if
			Set objNameSpace = objShell.Namespace(strFontSourcePath)
			Set objFolder = objFSO.getFolder(strFontSourcePath)
				For Each objFile In objFolder.filesz
					If LCase(right(objFile,4)) = ".ttf" OR LCase(right(objFile,4)) = ".otf" Then
						Set objFont = objNameSpace.ParseName(objFile.Name)
						If objFSO.FileExists("C:\WINDOWS\Fonts\" & objFile.Name) = False Then	'Check to see if font is already installed
							objFSO.CopyFile "\\US-SVR-NET-001\IDS Inc Shared\IT\Fonts\deploy\" & objFile.Name, "c:\windows\fonts\"	'copy font to local fonts folder
							if DisplayOutput = True then Wscript.Echo "Installed Font: " & objFont
							counter = counter + 1
							Set objFont = Nothing
						Else
						End If
					End If
				Next
		End If
		if DisplayOutput = True then
			Wscript.Echo ""
			Wscript.Echo "Complete. " & Counter & " fonts installed."								'Display completion message
		end if
		temp = WriteReg(RegKey, Installed, "REG_DWORD")			'Write registry key showing package has been run
	Else
		if DisplayOutput = True then Wscript.Echo "Font Source Path does not exist"
		wscript.quit
	End IF
End Function

Function ReadCompDetails()	'read computer name
	Set wshShell = WScript.CreateObject( "WScript.Shell" )
	strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )	
	if DisplayOutput = TRUE then
		Wscript.Echo "Computer name: " & strComputerName
		Wscript.Echo ""
	end if
	ReadCompDetails = strComputerName
End Function

Function SetFilePath()		'set font source based on location defined in computer name
	Dim strCompName, strCountry, strFilePath
	if DisplayOutput = TRUE then
		wscript.echo "Detecting Location"
		wscript.echo ""
		wscript.echo "Detecting Computer Name"
	end if
	strCompName = ReadCompDetails()
	if DisplayOutput = TRUE then wscript.echo "Detecting Location Based on Computer Name"
	strCountry = left(strCompName,2)
	if DisplayOutput = TRUE then wscript.echo "Location set as " & strCountry
	if strCountry = "US" then
		strFilePath = "\\US-SVR-NET-001\IDS Inc Shared\Marketing\Fonts\deploy\"
	else
		strFilePath = "\\UK-SVR-DFS-002\IDS Boldon\Marketing\Master Logos & Fonts\Fonts\deploy\"
	end if
	if DisplayOutput = TRUE then
		wscript.echo "Setting Font Source to: " & strFilePath
		wscript.echo ""
	end if
	SetFilePath = strFilePath
End Function