' ====================================================================
'
' Author	: Christophe Avonture
' Date		: October 2018
'
' Open all TXT files in the current folder and convert files from 
' Unix LF to Windows CRLF or the opposite
'
' Based on a script of Stephen Millard
' @see http://www.thoughtasylum.com/blog/2015/6/28/vbscript-flip-line-endings.html
'
' ====================================================================

Option Explicit

' Set to True to run without any message box notifications
Const SILENT_MODE = false

' --------------------------------------------------------
' Changes Windows line endings (CRLF) to UNIX (LF) or
' the opposite, depending on the sType parameter
'
' When sType is CRLF, if the processed file has LF, convert
' otherwise skip since the file is already correct.
'
' When sType is LF, if the processed file has CRLF, convert
' otherwise skip since the file is already correct.
' --------------------------------------------------------
Function ConvertLineEndings(sFileName, sType)

    Const FOR_READING = 1
    Const FOR_WRITING = 2

    Dim sContent, sMessage
    Dim objFSO, objFile
    Dim bContinue

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	' Make sure the file exists
	If objFSO.FileExists(sFileName) then

		' Get file content; read the file entirely
		Set objFile = objFSO.OpenTextFile(sFileName, FOR_READING)
		sContent = objFile.ReadAll
		objFile.Close

		' Initialize
		sMessage = "Lines already ends with " & sType & ", skip the file"

		' Detect the current, used, line end char LF or CRLF?
		If InStr(1, sContent, vbCrLf, vbTextCompare) > 0 Then
			' Convert only if the target type is LF
			bContinue = (sType = "LF")

			If bContinue Then 
				sMessage = "Windows (CRLF) to UNIX (LF)."
				sContent = Replace(sContent, vbCrLf, vbLf)
			End if

		Else 

			' Convert only if the target type is LF
			bContinue = (sType = "CRLF")

			If bContinue Then 
				sMessage = "UNIX (LF) to Windows (CRLF)."
				sContent = Replace(sContent, vbLf, vbCrLf)
			End if
			
		End If

		If bContinue Then 

			' Delete original file
			objFSO.DeleteFile sFileName

			' Write the file with the converted line endings
			Set objFile = objFSO.OpenTextFile(sFileName, FOR_WRITING, True, False)
			objFile.Write(sContent)

		End If

		ConvertLineEndings = True

	Else

		' Should never comes here
		sMessage = "File Not Found: " & sFileName
		ConvertLineEndings = False

	End If

	If Not SILENT_MODE then 
		wScript.echo "      " & sMessage
	End if

End Function

Sub ShowHelp()

	wScript.echo " =========================================="
	wScript.echo " = Convert from LF to CRLF and vice-versa ="
	wScript.echo " =========================================="
	wScript.echo ""
	wScript.echo " Please specify: "
	wScript.echo "   1. the folder name (or just a '.' for the current folder)"
	wScript.echo "   2. the extension of files to process (f.i. 'txt')"
	wScript.echo "   3. the desired line ending ('LF' or 'CRLF')"
	wScript.echo ""
	wScript.echo " Examples: " 
    wScript.echo "      cscript " & Wscript.ScriptName & " . txt CRLF"
    wScript.echo "      cscript " & Wscript.ScriptName & " C:\Data txt CRLF"
	wScript.echo ""

	wScript.echo " To get more info, please read https://github.com/cavo789/vbs_convert_LF_CRLF"
	wScript.echo ""

	wScript.quit

End sub

' -------------------
' --- ENTRY POINT ---
' -------------------

Dim objFSO, objFolder, objFiles, objFile
Dim sPath, sExtension, sType

	' We need three arguments
	If (wScript.Arguments.Count < 3) Then

		Call ShowHelp

	Else

		Set objFSO = CreateObject("Scripting.FileSystemObject")
		
		' Get the path specified on the command line (f.i. "." for the current folder)
		sPath = trim(wScript.Arguments.Item(0))
		
		If (sPath = ".") Then
			' Get the current folder
			sPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(wScript.ScriptFullName)
		End If

		' Get the file extension to process (f.i. "txt")
		sExtension = Trim(wScript.Arguments.Item(1))

		' Get the desired type (LF or CRLF)
		sType = UCase(Trim(wScript.Arguments.Item(2)))

		If (sType <> "LF") And (sType <> "CRLF") Then
			wScript.echo " ERROR - Invalid type. Should be LF or CRLF and nothing else"
  
		Else

			' Ok, we can process file(s)
			If Not SILENT_MODE then
				wScript.echo "Process every ." & sExtension & " files under " & sPath 
				wScript.echo "Convert to " & sType & " when needed"
				wScript.echo ""
			End If

			' ----------------------------------
			' Loop and retrieve every .txt files
			Set objFolder = objFSO.GetFolder(sPath)

			Set objFiles = objFolder.Files

			For Each objFile in objFiles
				' Case insensitive comparaison
				If StrComp(objFSO.GetExtensionName(objFile.name), sExtension, vbTextCompare) = 0 Then

					If Not SILENT_MODE then 
						wScript.echo "   Process " & objFile.Name
					End if

					ConvertLineEndings objFile.Name, sType

				End If
			Next

		End If

		
	End If
