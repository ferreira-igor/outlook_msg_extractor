'Title: Outlook Msg Extractor
'Description: Extrai anexos de arquivos ".msg" do Outlook.
'Author: Igor Ferreira
'Version: 1.1.0

Dim ol_obj, fs_obj, folder_path, mail_file, mail_name, f, i

Set ol_obj = CreateObject("Outlook.Application")
Set fs_obj = CreateObject("Scripting.FileSystemObject")

folder_path = fs_obj.GetParentFolderName(WScript.ScriptFullName)

For Each f In fs_obj.GetFolder(folder_path).Files
    If LCase(fs_obj.GetExtensionName(f)) = "msg" Then
        Set mail_file = ol_obj.CreateItemFromTemplate(f.Path)
        mail_name = Left(fs_obj.GetFileName(f), (InStrRev(fs_obj.GetFileName(f), ".", -1, vbTextCompare) - 1))
        fs_obj.CreateFolder(mail_name)
        If mail_file.Attachments.Count > 0 Then
            For i = 1 To mail_file.Attachments.Count
                mail_file.Attachments(i).SaveAsFile folder_path & "\" & mail_name & "\" & mail_file.Attachments(i).FileName
            Next
        End If
    End If
Next

MsgBox "Ok!"
