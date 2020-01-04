Attribute VB_Name = "xlsmToxlam"
Option Explicit

Sub アドイン保存()

    'アドインに変換するファイルを取得
    Dim vFileName As Variant
    vFileName = Application.GetOpenFilename("Excelファイル,*.x*")
    If vFileName = False Then
        Exit Sub
    End If
    
    On Error Resume Next
    Dim objBook As Workbook
    Application.EnableEvents = False 'マクロの実行を抑制する
    Set objBook = Workbooks.Open(vFileName, 0, True) 'Linkの更新をしない、ReadOnly
    Application.EnableEvents = True
    If objBook Is Nothing Then
        Call MsgBox(vFileName & vbCrLf & " が開けませんでした。")
        Exit Sub
    End If
    On Error GoTo 0
    
    'アドインの時は設定要
    objBook.IsAddin = True

'    'フォルダ名とファイル名を取得
'    Dim strFolder As String 'フォルダ名
'    Dim strFile   As String 'ファイル名(拡張子除く)
'    With CreateObject("Scripting.FileSystemObject")
'        strFolder = .GetParentFolderName(vFileName)
'        strFile = .GetBaseName(vFileName)
'    End With
    
    '拡張子を.xlamに置換
    Dim strAddinFile As String
    With CreateObject("Scripting.FileSystemObject")
        strAddinFile = Replace(vFileName, "." & .GetExtensionName(vFileName), ".xlam")
    End With
    
    Dim strPassword As String
    strPassword = InputBox("設定するパスワードを入力してください")
    
    Call objBook.SaveAs(strAddinFile, xlOpenXMLAddIn, strPassword)
    Call objBook.Close(False) '保存せずに閉じる
    
    Call MsgBox("個人情報の削除　と" & vbLf & "バージョン番号の変更を忘れないでください")
End Sub
