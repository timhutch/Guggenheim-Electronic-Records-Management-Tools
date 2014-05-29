' pres&access.vbs (VBScript)
' DOC, PPT, DOCX, PPTX, WPD, MLM, XLS, RTF normalization (access & preservation)
' for Archivematica


' Copyright (c) 2014 The Solomon R. Guggenheim Foundation. 

 
' Permission is hereby granted, free of charge, to any person obtaining a copy of this 
' software (the “Software”), to copy, modify, publish, distribute and otherwise use the 
' Software for his/her personal use, provided that, such user shall include the above 
' copyright line and this permission notice in all copies of the Software.
 
' ***************
 
' Access to the Software is provided at the Guggenheim’s discretion.  The Guggenheim may 
' terminate access to the Software at any time without prior notice.  Use of the Software 
' is entirely at the user’s risk.  The Guggenheim is not responsible for any damage, loss, 
' or theft that may occur as a result of the user’s use of the Software.
 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR 
' PURPOSE AND NON-INFRINGEMENT.  IN NO EVENT SHALL THE COPYRIGHT HOLDER BE LIABLE FOR ANY 
' CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, 
' ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR ITS USE.
 
' As of May 1, 2014

' Written by Anthony Cocciolo, Electronic Records Consultant


Option Explicit

Dim zSourceDir
Dim oWord, FSO, masterCount, oPPT, oExcel, oBook, ext, folder, stringi, error_log, ofile
Dim manualNorm


Set FSO = CreateObject("Scripting.FileSystemObject")

' Need to provide the source directory
zSourceDir = InputBox ("Enter the directory with WordPerfect, MLM, RTF, MS Word, MS Excel and/or MS Powerpoint files are contained (e.g., c:\wpfiles).  Please note that this program will also check sub-folders within the folder you specify, and ignore files not in these formats:")

if NOT fso.FolderExists (zSourceDir) then
	MsgBox ("Folder does not exist.  Quitting")
	WScript.Quit
else
	MsgBox ("Press OK and the process will begin.  This may take awhile.  You will be notified when the process is complete")
end if

' standardize path
Set folder = fso.GetFolder(zSourceDir)
zSourceDir = folder.Path 


Set oWord = CreateObject("Word.Application")

masterCount = 0
error_log = ""

ConvFiles (zSourceDir)


' creates manual formalization folders
function createManualPaths (path)

	manualNorm = path & "\manualNormalization"
	if not FSO.folderExists(manualNorm) then
		FSO.createFolder (manualNorm)
	end if
	
	if not FSO.folderExists (manualNorm & "\access") then
		FSO.createFolder (manualNorm & "\access")
	end if
	
	if not FSO.folderExists (manualNorm & "\preservation") then
		FSO.createFolder (manualNorm & "\preservation")
	end if

	createManualPaths = manualNorm
end function 

' recursively normalize files for preservation and access
sub ConvFiles (currentDir)
	Dim oFolder, f, oDoc, manualNorm, subfolders, sf, frontPath
	
	Set oFolder = FSO.GetFolder(currentDir)
	
	' normalize all subfolders recursively
	Set subfolders = oFolder.SubFolders
   	For Each sf in subfolders
   		
   		if sf.name <> "manualNormalization" and not (sf.path = currentDir) then
    		ConvFiles (sf.Path)
    	end if
    Next
	
	' normalize each file
	for each f in oFolder.Files
	
		' ignore files that are hidden or system files
		if (   (f.attributes and 2) OR (f.attributes AND 4)  ) then
			' do nothing for hidden or system files
		else
	
		
			ext = lcase(FSO.GetExtensionName(f.path))
			
			' only apply to valid files
			if (ext = "wpd" Or ext = "mlm" or ext = "doc" or ext = "docx" or ext = "xls" or ext = "ppt" or ext = "pptx" or ext="rtf") then
					
				manualNorm = createManualPaths (zSourceDir)
				
				stringi = len(f.parentfolder) - len(zSourceDir) - 1
				if stringi < 0 then 
					stringi = 0
				end if
				
				frontPath = right (f.ParentFolder, stringi )
				if ( Not (frontPath = "")) then
					frontPath = frontPath & "\"
					
				end if 
				
				subCreateFolders manualNorm & "\preservation\" & frontPath
				subCreateFolders manualNorm & "\access\" & frontPath
				
				
				if ext  = "wpd" OR ext = "mlm" OR ext  = "doc" or ext = "docx" or ext = "rtf" then
				
					On Error Resume Next
					Set oDoc = oWord.Documents.Open(f.path, , True)
					
				
					if Err.number <> 0 then
						Error_log = Error_log & "Unable to open: " & f.path & " (Description: " & err.description & ")" & vbNewline 
						Set oDoc = Nothing
						
					else
						Set oDoc = oWord.ActiveDocument
						
						' create docx - skip if already DOCX or RTF
						if NOT (ext = "rtf" or ext = "docx") then				
							
							oDoc.SaveAs manualNorm  & "\preservation\" & frontPath  & f.name & ".docx", 16 
				
							if Err.Number <> 0 then
								Error_log = Error_log & "Error creating: " & manualNorm  & "\preservation\" & frontPath  & f.name & ".docx" &  " (Description: " & err.description & ")" & vbNewline
							end if
					
							
						end if
				
						
						' create pdf
						oDoc.SaveAs manualNorm  & "\access\" & frontPath  & f.name & ".pdf", 17
					
						if Err.Number <> 0 then
							Error_log = Error_log & "Error creating: " & manualNorm  & "\access\" & frontPath  & f.name & ".pdf" & " (Description: " & err.description & ")" & vbNewline
						end if
						
						oDoc.Close 0
						Set oDoc = Nothing
					end if
				
					On Error GoTo 0
					
				else if ext = "xls" then
					Set oExcel = CreateObject("Excel.Application")
					
					
					On Error Resume Next
					Set oBook = oExcel.Workbooks.Open (f.path)
					if Err.Number <> 0 then
						Error_log = Error_log & "Unable to open: " & f.path  & " (Description: " & err.description & ")" & vbNewline 
					else
					
						oBook.SaveAs  manualNorm  & "\preservation\" & frontPath & f.name & ".xlsx", 51
						if Err.Number <> 0 then
							Error_log = Error_log & "Error creating: " &  manualNorm  & "\preservation\" & frontPath & f.name & ".xlsx" & " (Description: " & err.description & ")" & vbNewline
						end if
						oBook.SaveAs  manualNorm  & "\access\" & frontPath & f.name & ".xlsx",51					
						if Err.Number <> 0 then
							Error_log = Error_log & "Error creating: " &  manualNorm  & "\access\" & frontPath & f.name & ".xlsx" & " (Description: " & err.description & ")" & vbNewline
						end if
		

						oBook.Close 0
					
					end if
					On Error Goto 0
					
					Set oBook = Nothing
					oExcel.Quit
					Set oExcel = Nothing
			
				else if (ext = "ppt" or ext = "pptx") then
			
					
					Set oPPT = CreateObject("Powerpoint.Application")
					oPPT.visible = true
				
					On Error Resume Next 
					Set oPPT = oPPT.Presentations.Open(f.path, , True)
					
					if Err.Number <> 0 then
						Error_log = Error_log & "Unable to open: " & f.path & " (Description: " & err.description & ")" & vbNewline 
					else
					
						' create pdf
						oPPT.SaveAs manualNorm  & "\access\" & frontPath & f.name & ".pdf", 32
						
						if Err.Number <> 0 then
							Error_log = Error_log & "Error creating: " & manualNorm  & "\access\" & frontPath & f.name & ".pdf" & " (Description: " & err.description & ")" & vbNewline
						end if
					
						if (NOT ext = "pptx") then ' don't create PPTX if already PPTX
							oPPT.SaveAs manualNorm & "\preservation\" & frontPath  & f.name & ".pptx", 24
							if Err.Number <> 0 then
								Error_log = Error_log & "Error creating: " & manualNorm & "\preservation\" & frontPath  & f.name & ".pptx" & " (Description: " & err.description & ")" & vbNewline
							end if
						end if
						oPPT.Close 
					end if 
					
					On Error GoTo 0
		
					
					Set oPPT = Nothing
				end if
				end if
				end if
				
				masterCount = masterCount + 1
				
			end if
			
		end if

	next 

end sub

Sub subCreateFolders(strPath)
   Dim objFileSys
   Dim strNewFolder
    
   Set objFileSys = CreateObject("Scripting.FileSystemObject")

   If Right(strPath, 1) <> "\" Then
      strPath = strPath & "\"
   End If

   strNewFolder = ""
   Do Until strPath = strNewFolder
      strNewFolder = Left(strPath, InStr(Len(strNewFolder) + 1, strPath, "\"))
    
      If objFileSys.FolderExists(strNewFolder) = False Then
         objFileSys.CreateFolder(strNewFolder)
      End If
   Loop
End Sub


oWord.Quit
Set oWord = Nothing


if error_log = "" then
	MsgBox "Creation of normalized files are completed.  " & masterCount & " files were normalized."
else
	Set ofile = fso.OpenTextFile ("pres_and_access.log", 2, true)
	ofile.writeline error_log
	ofile.close
	Set ofile = Nothing
	Set fso = Nothing

	MsgBox masterCount & " Files were attempted to be normalized, but there was an error with one or more.  These are included in the log file: pres_and_access.log"
end if







