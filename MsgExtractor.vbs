'Title: Outlook Msg Extractor
'Description: Extracts attachments from Outlook msg files.
'Author: Igor Ferreira
'Version: 1.0.0

Dim ol_obj, fs_obj, folder_path, mail_file, f, i

Set ol_obj = CreateObject("Outlook.Application")
Set fs_obj = CreateObject("Scripting.FileSystemObject")

folder_path = fs_obj.GetParentFolderName(WScript.ScriptFullName)

For Each f In fs_obj.GetFolder(folder_path).Files
    If LCase(fs_obj.GetExtensionName(f)) = "msg" Then
        Set mail_file = ol_obj.CreateItemFromTemplate(f.Path)
        If mail_file.Attachments.Count > 0 Then
            For i = 1 To mail_file.Attachments.Count
                mail_file.Attachments(i).SaveAsFile folder_path & "\" & mail_file.Attachments(i).FileName
            Next
        End If
    End If
Next

MsgBox "Successfully extracted attachments!"