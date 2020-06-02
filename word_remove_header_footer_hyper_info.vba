 '此代码功能，批量为选择的WORD文件进行 删除页眉页脚超链接作者等内容的操作，快捷键 ALT-F11 进入VBA
 'Author: csoftcn
 'Time: May, 2020
 '*****************
Sub remove_word_header_footer_hyper_info()
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
                For Each oSec In oDoc.Sections '文档的节中循环
                    Set myRange = oSec.Headers(wdHeaderFooterFirstPage).Range
                    myRange.Delete '删除页眉中的内容
                    Set myRange = oSec.Headers(wdHeaderFooterPrimary).Range
                    myRange.Delete '删除页眉中的内容
                    Set myRange = oSec.Footers(wdHeaderFooterFirstPage).Range
                    myRange.Delete '删除页脚中的内容
                    Set myRange = oSec.Footers(wdHeaderFooterPrimary).Range
                    myRange.Delete '删除页脚中的内容
                Next

                For i = 1 To oDoc.Hyperlinks.Count  '删除所有超链接
                    oDoc.Hyperlinks(i).Delete
                Next

                With oDoc.BuiltInDocumentProperties '删除所有文档信息
                    .Item("title").Value = ""
                    .Item("subject").Value = ""
                    .Item("author").Value = ""
                    .Item("manager").Value = ""
                    .Item("company").Value = ""
                    .Item("comments").Value = ""
                    .Item("keywords").Value = ""
                    .Item("category").Value = ""
                    .Item("timelastsaved").Value = ""
                End With

                oDoc.Close True
            Next
        End If
    End With
End Sub

