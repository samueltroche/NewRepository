Dim FSO,DWName, FileTXT, Linea, Total, FileOutput, ObjStr
Set FSO = CreateObject("Scripting.FileSystemObject")
'------------------------------------------------------
InputFOX1 = "C:\RICOH\Repository\DocuWare\F1_fox"
InputFOX4 = "C:\RICOH\Repository\DocuWare\F4_foxlm"
OutputFOX = "C:\RICOH\Repository\DocuWare\Input_fox"
InputDHS = "C:\RICOH\Repository\DocuWare\F2_dhs"
OutputDHS = "C:\RICOH\Repository\DocuWare\Input_dhs"
InputANGLOLAB = "C:\RICOH\Repository\DocuWare\F3_anglolab"
OutputANGLOLAB = "C:\RICOH\Repository\DocuWare\Input_anglolab"
InputINVOICEFOX = "C:\RICOH\Repository\DocuWare\F5_InvoiceFOX"
OutputINVOICEFOX = "C:\RICOH\Repository\DocuWare\Input_InvoiceFox"
InputBuild = "C:\RICOH\Repository\DocuWare\InputBuild"
OutputBuild = "C:\RICOH\Repository\DocuWare\OutputBuild"
'------------------------------------------------------
'Call GetFiles(InputFOX1, OutputFOX, ".txt","C:\RICOH\Repository\Docuware\backup\backup-fox\")
'Call GetFiles(InputFOX4, OutputFOX, ".txt","C:\RICOH\Repository\Docuware\backup\backup-fox\")
'Call GetFilesPDF(InputDHS, OutputDHS, ".pdf","C:\RICOH\Repository\Docuware\backup\backup-dhs\")
'Call GetFilesPDF(InputANGLOLAB, OutputANGLOLAB, ".pdf", "C:\RICOH\Repository\Docuware\backup\backup-anglolab\")
'Call GetFiles(InputINVOICEFOX, OutputINVOICEFOX, ".txt", "C:\RICOH\Repository\Docuware\backup\backup-invoicefox\")
Call RenameBuildXml(InputBuild, OutputBuild, ".pdf")
Wscript.Echo "End Process"

Function GetFiles(InputName, OutputName, FT, BAK)
	On Error Resume Next
	Dim ObjFolder, ObjSubFolders, ObjFiles, ObjFile
	Set ObjFolder = FSO.GetFolder(InputName)
	Set ObjFiles = ObjFolder.Files
	Set ObjStr = CreateObject("ADODB.Stream")
	For Each ObjFile In ObjFiles
		Total = ""
		DWName = Mid(ObjFile.Name,1,Len(ObjFile.Name)-4)
		DWName1 = Mid(ObjFile.Name,1,Len(ObjFile.Name)-4)'cambio 10-02-2021
		If Right(LCase(ObjFile.Name),4) = ".xml" And FSO.FileExists(InputName & "\" & DWName & FT) Then
			Set FileTXT = FSO.OpenTextFile(objFile.Path,1)
			Do While NOT FileTXT.AtEndOfStream
				Linea = FileTXT.ReadLine
				Total= Total + Linea
			Loop
			FSO.MoveFile InputName & "\" & DWName & FT, OutputName & "\" & DWName & FT
			ObjStr.CharSet = "utf-8"
			ObjStr.Open
			ObjStr.WriteText Total
			ObjStr.SaveToFile OutputName & "\" & DWName & ".dwcontrol", 2
			ObjStr.Close
			FileTXT.Close
			FSO.MoveFile InputName & "\" & DWName1 & ".xml", BAK & DWName1 & ".xml" '2021-05-19
'FSO.DeleteFile InputName & "\" & DWName1 & ".xml" 'cambio 2021-02-10
		End If
	Next
End Function

Function GetFilesPDF(InputName, OutputName, FT, BAK)
	On Error Resume Next
	Dim ObjFolder, ObjSubFolders, ObjFiles, ObjFile
	Set ObjFolder = FSO.GetFolder(InputName)
	Set ObjFiles = ObjFolder.Files
	For Each ObjFile In ObjFiles
		Total = ""
		DWName = Mid(ObjFile.Name,1,Len(ObjFile.Name)-4)
		DWName1 = Mid(ObjFile.Name,1,Len(ObjFile.Name)-4)'cambio 10-02-2021
		If Right(LCase(ObjFile.Name),4) = ".xml" And FSO.FileExists(InputName & "\" & DWName & FT) Then
			Set FileTXT = FSO.OpenTextFile(objFile.Path,1)
			Do While NOT FileTXT.AtEndOfStream
				Linea = FileTXT.ReadLine
				Total= Total + Linea
				If Right(LCase(Linea),20) = "xmlschema-instance"">" Then
'nothing
				End If
			Loop
			Set FileOutput = FSO.CreateTextFile(OutputName & "\" & DWName & ".dwcontrol",True)
			FSO.MoveFile InputName & "\" & DWName & FT, OutputName & "\" & DWName & FT
			FileOutput.Write Total
			FileOutput.Close
			FileTXT.Close
'FSO.DeleteFile Input & "\" & DWName1 & ".xml" 'cambio 10-02-2021
			FSO.MoveFile InputName & "\" & DWName1 & ".xml", BAK & DWName1 & ".xml" 'cambio 10-02-2021
		End If
	Next
End Function

Function RenameBuildXml(InputBuild, OutputBuild, TypeDoc)
'On Error Resume Next
	Dim ObjFolder, ObjSubFolders, ObjFiles, ObjFile, FileArr
	Set ObjFolder = fso.GetFolder(InputBuild)
	Set ObjFiles  = ObjFolder.Files
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	Set ObjStr = CreateObject("ADODB.Stream")
	For Each ObjFile In ObjFiles
		''wscript.echo ObjFile.Name
		If (InStr(objFile.Name, "-" ) > 0) Then
			FileArr = split(objFile.Name,"-")
			If (UBound(FileArr) >= 2 And Len(Trim(FileArr(0)))=3 And Len(Trim(FileArr(1)))=8) Then
				With New RegExp
					.Pattern = "^(L|M|N|C|H|U)[0-9]{7}$"
					.IgnoreCase = True
					.Global = True
					Set Matches = .Execute(FileArr(1))
				End With
				If Matches.Count = 1 Then
					Set CsvFile   = ObjFSO.OpenTextFile("C:\RICOH\Repository\DocuWare\MasterData.csv")
					Do While CsvFile.AtEndOfStream <> True
						CsvArr = Split(CsvFile.ReadLine, ",")
						If (Trim(CsvArr(0)) = Trim(FileArr(0))) Then
							''wscript.echo Trim(CsvArr(1)) 'Se debe colocar el codigo a hacer con las 2 variables
							''creando archivo
							Total = ""
							Set FileTXT = FSO.OpenTextFile("C:\RICOH\Repository\DocuWare\PlantillaDwcontrol.txt")
								Do While NOT FileTXT.AtEndOfStream
								Linea = FileTXT.ReadLine
								Total= Total + Replace(Replace(Linea,"VRecord",FileArr(1)),"VDocumentType",CsvArr(1))
								' Total= Total + Replace(Linea,"VDocumentType",CsvArr(1))
							Loop
							ObjStr.CharSet = "utf-8"
							ObjStr.Open
							ObjStr.WriteText Total
							objStr.SaveToFile OutputBuild & "\" & Replace(objFile.Name,".pdf",".dwcontrol"), 2
							ObjStr.Close
							''FSO.MoveFile InputBuild & "\" & objFile.Name, OutputBuild & "\" & objFile.Name
							FSO.CopyFile InputBuild & "\" & objFile.Name, OutputBuild & "\" & objFile.Name, 1
							wscript.echo "Stop"
							FSO.DeleteFile InputBuild & "\" & objFile.Name
							''FileTXT.Close
						End If
					Loop
					CsvFile.Close
				End If
			End If
		End if
	Next
	Set ObjSubFolders = ObjFolder.SubFolders
	For Each ObjFolder In ObjSubFolders
		Call RenameBuildXml(ObjFolder.Path, "", "")
	Next
End Function