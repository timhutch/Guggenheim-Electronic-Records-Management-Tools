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
Dim oWord, FSO, masterCount, oPPT, oExcel, oBook, ext, folder, stringi 
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
				
				
				if ext  = "wpd" OR ext = "mlm" OR ext  = "doc"  then
				
					Set oDoc = oWord.Documents.Open(f.path, , True)
					Set oDoc = oWord.ActiveDocument
				
				
					' create docx
				
					oDoc.SaveAs manualNorm  & "\preservation\" & frontPath  & f.name & ".docx", 16 
				
					' create pdf
					oDoc.SaveAs manualNorm  & "\access\" & frontPath  & f.name & ".pdf", 17
				
					
		
					oDoc.Close 0
					Set oDoc = Nothing
				else if (ext = "docx" OR ext = "rtf") then
				
			
					'on error resume next
					'FSO.CopyFile f.path, manualNorm  & "\preservation\" & frontPath  & f.name, true
					'on error goto 0
				
				
					Set oDoc = oWord.Documents.Open(f.path, , True)
					Set oDoc = oWord.ActiveDocument			
				
					' create pdf
					oDoc.SaveAs manualNorm  & "\access\" & frontPath & f.name & ".pdf", 17
				
		
					oDoc.Close 0
					Set oDoc = Nothing

				else if ext = "xls" then
					Set oExcel = CreateObject("Excel.Application")
					
					Set oBook = oExcel.Workbooks.Open (f.path)
					
					oBook.SaveAs  manualNorm  & "\preservation\" & frontPath & f.name & ".xlsx", 51
					oBook.SaveAs  manualNorm  & "\access\" & frontPath & f.name & ".xlsx",51					
		
					oBook.Close 0
					Set oBook = Nothing
					oExcel.Quit
					Set oExcel = Nothing
			
				else if (ext = "ppt") then
			
					Set oPPT = CreateObject("Powerpoint.Application")
					oPPT.visible = true
				
					Set oPPT = oPPT.Presentations.Open(f.path, , True)
					
				
					' create pdf
					oPPT.SaveAs manualNorm  & "\access\" & frontPath & f.name & ".pdf", 32
					oPPT.SaveAs manualNorm & "\preservation\" & frontPath  & f.name & ".pptx", 24
				
		
					oPPT.Close 
					
					Set oPPT = Nothing
				else if (lcase(FSO.GetExtensionName(f.path)) = "pptx") then
			
					Set oPPT = CreateObject("Powerpoint.Application")
					oPPT.visible = true

				
					'on error resume next
					'FSO.CopyFile f.path, manualNorm  & "\preservation\" & frontPath  & f.name, true
					'on error goto 0
				
					Set oPPT = oPPT.Presentations.Open(f.path, , True)
					
				
					' create pdf
					oPPT.SaveAs manualNorm  & "\access\" & frontPath & f.name & ".pdf", 32
				
		
					oPPT.Close 
					
					Set oPPT = Nothing
			
				end if
				end if
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


Set FSO = Nothing
oWord.Quit
Set oWord = Nothing

MsgBox "Creation of normalized files are completed.  " & masterCount & " files were normalized."







