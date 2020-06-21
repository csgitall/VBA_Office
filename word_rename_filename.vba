Sub rename_by_firstSentence()    '此代码功能，用文档第一句话作为文件名另存为新文档

    Dim myDialog As FileDialog, oDoc As Document, oSec As Section
    Dim oFile As Variant, myRange As Range
    On Error Resume Next
    '定义一个文件夹选取对话框
    Set myDialog = Application.FileDialog(msoFileDialogFilePicker)
    With myDialog
        .Filters.Clear    '清除所有文件筛选器中的项目
        .Filters.Add "所有 WORD 文件", "*.doc,*.docx", 1    '增加筛选器的项目为所有WORD文件
        .AllowMultiSelect = True    '允许多项选择
        If .Show = -1 Then    '确定
            For Each oFile In .SelectedItems    '在所有选取项目中循环
                Set oDoc = Word.Documents.Open(FileName:=oFile, Visible:=False)
                newFilename = oDoc.Sentences.Item(1).Text '找到第一个句子，并去除回车和换行符
                newFilename = Replace(newFilename, Chr(10), "")
                newFilename = Replace(newFilename, Chr(13), "")
				newFilename = Replace(newFilename,  " ", "")
                'MsgBox oDoc.Path
                oDoc.SaveAs2 FileName:=oDoc.Path & "\" & newFilename
                oDoc.Close True
                'If newFilename <> "" Then
                '    oldFilename = oFile.Name
                '    Name oFile As newFilename
                'End If
            Next
        End If
    End With
End Sub

