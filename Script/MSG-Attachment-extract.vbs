'Variables
Dim ol, fso, folderPath, destPath, f, msg, i
'Loading objects
Set ol  = CreateObject("Outlook.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
'Setting MSG files path
folderPath = fso.GetParentFolderName(WScript.ScriptFullName)
'Setting destination path
destPath = folderPath	'* I am using the same 
WScript.Echo "==> "& folderPath
'Looping for files
For Each f In fso.GetFolder(folderPath).Files
	'Filtering only MSG files
	If LCase(fso.GetExtensionName(f)) = "msg" Then
		'Opening the file
		Set msg = ol.CreateItemFromTemplate(f.Path)
		'Checking if there are attachments
		If msg.Attachments.Count > 0 Then
			'Looping for attachments
			For i = 1 To msg.Attachments.Count
				'Checking if is a Excel file
				If LCase(Mid(msg.Attachments(i).FileName, InStrRev(msg.Attachments(i).FileName, ".") + 1 , 3)) = "xls" Then
					WScript.Echo f.Name &" -> "& msg.Attachments(i).FileName
					'Saving the attachment
					msg.Attachments(i).SaveAsFile destPath &"\"& msg.Attachments(i).FileName
				End If
			Next
		End If
	End If
Next
MsgBox "Attachments successfully extracted!"& vbcrlf &"Anexos extraídos com sucesso!"
