 '此代码功能，批量为选择的PowerPoint文件进行 删除超链接作者等内容的操作，快捷键 ALT-F11 进入VBA
  '增加功能，批量删除 备注文字
 'Author: csoftcn
 'Update:  June, 2020
 'Time: May, 2020
 '*****************
Sub remove_ppt_hyper_info()
    Dim myDialog As FileDialog
    Dim oPpt As Presentation
    Dim oFile As Variant
    On Error Resume Next
    '定义一个文件夹选取对话框
    Set myDialog = Application.FileDialog(msoFileDialogFilePicker)
    With myDialog
        .Filters.Clear    '清除所有文件筛选器中的项目
        .Filters.Add "所有 PowerPoint 文件", "*.ppt,*.pptx", 1    '增加筛选器的项目为所有PowerPoint文件
        .AllowMultiSelect = True    '允许多项选择
        If .Show = -1 Then    '确定
            For Each oFile In .SelectedItems    '在所有选取项目中循环
                Set oPpt = Application.Presentations.Open(FileName:=oFile, WithWindow:=msoFalse)

                For Each hl In oPpt.Slides.Range.Hyperlinks '删除所有超链接
                    hl.Delete
                Next

                With oPpt.BuiltInDocumentProperties '删除所有文档信息
                    .Item("Title").Value = ""
                    .Item("Subject").Value = ""
                    .Item("Author").Value = ""
                    .Item("Last author").Value = ""
                    .Item("manager").Value = ""
                    .Item("Company").Value = ""
                    .Item("Comments").Value = ""
                    .Item("keywords").Value = ""
                    .Item("Category").Value = ""
                    .Item("Last save time").Value = ""
                End With

                For Each sld In oPpt.Slides '删除所有备注
                    sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange = ""
                Next sld


                oPpt.Save
                oPpt.Close
            Next
        End If
    End With
End Sub
