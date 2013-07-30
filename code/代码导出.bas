Attribute VB_Name = "代码导出"
' https://gist.github.com/rhiguchi/913153
Option Explicit
 
Private exportSelf As Boolean
Public Const MODULE_NAME_SPACE As String = "VBACodeExporter"
 
Private Enum ComponentType
    STANDARD_MODULE = 1
    CLASS_MODULE = 2
    USER_FORM = 3
    DOCUMENT_MODULE = 100
End Enum
 
' VBAコードを出力する。保存場所はダイアログで出力先を尋ねる。
' 出力先に同じファイルがあった場合上書きされる
Public Sub 导出代码() 'exportVBACodesWithFileDialog
    Dim destDir As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ThisWorkbook.path
        .ButtonName = "OK"
        If .Show = 0 Then Exit Sub
    
        destDir = .SelectedItems(1)
    End With
    
    ExportVBACodes destDir
End Sub
 
' すべてのVBAコードを出力
' 出力先に同じファイルがあった場合上書きされる
' @param destDir 出力先ディレクトリ
' @param wb 出力するワークブック。指定しない場合、ThisWorkbookとなる
Public Sub ExportVBACodes(destDir As String, Optional wb As Workbook = Nothing)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Dim vbcomp As Object
    For Each vbcomp In wb.VBProject.VBComponents
        Dim fileName As String: fileName = vbcomp.Name
        
        ' ソースの種類ごとに拡張子を決める
        ' 出力しない場合、ファイル名は空になる
        Select Case vbcomp.Type
            Case STANDARD_MODULE
                fileName = fileName & ".bas"
            Case CLASS_MODULE, DOCUMENT_MODULE
                fileName = fileName & ".cls"
            Case USER_FORM
                fileName = fileName & ".frm"
            Case Else
                fileName = ""
        End Select
        
        ' 自身のネームスペースのファイルは出力しない
        If Not exportSelf And InStr(1, fileName, MODULE_NAME_SPACE) = 1 Then
            fileName = ""
        End If
        
        ' ファイル名が空でなければUTF8で出力を行う
        If fileName <> "" Then
            Dim filePath As String: filePath = destDir & "\" & fileName
            vbcomp.Export filePath
            convertCharCode_SJIS_to_utf8 filePath
        End If
    Next
End Sub
 
' ファイルの文字コードをSJISからUTF8(BOM無し)に変換する
Private Sub convertCharCode_SJIS_to_utf8(file As String)
    Dim destWithBOM As Object: Set destWithBOM = CreateObject("ADODB.Stream")
    With destWithBOM
        .Type = 2
        .Charset = "utf-8"
        .Open
        
        ' ファイルをSJIS で開いて、dest へ 出力
        With CreateObject("ADODB.Stream")
            .Type = 2
            .Charset = "gb2312" 'shift-jis
            .Open
            .LoadFromFile file
            .Position = 0
            .copyTo destWithBOM
            .Close
        End With
        
        ' BOM消去
        ' 3バイト無視してからバイナリとして出力
        .Position = 0
        .Type = 1 ' adTypeBinary
        .Position = 3
        
        Dim dest: Set dest = CreateObject("ADODB.Stream")
        With dest
            .Type = 1 ' adTypeBinary
            .Open
            destWithBOM.copyTo dest
            .savetofile file, 2
            .Close
        End With
        
        .Close
    End With
End Sub
 
 
Public Property Get isExportSelf() As Boolean
    isExportSelf = exportSelf
End Property
 
Public Property Let isExportSelf(ByVal vNewValue As Boolean)
    exportSelf = vNewValue
End Property
