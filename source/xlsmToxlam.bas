Attribute VB_Name = "xlsmToxlam"
Option Explicit

Sub �A�h�C���ۑ�()

    '�A�h�C���ɕϊ�����t�@�C�����擾
    Dim vFileName As Variant
    vFileName = Application.GetOpenFilename("Excel�t�@�C��,*.x*")
    If vFileName = False Then
        Exit Sub
    End If
    
    On Error Resume Next
    Dim objBook As Workbook
    Application.EnableEvents = False '�}�N���̎��s��}������
    Set objBook = Workbooks.Open(vFileName, 0, True) 'Link�̍X�V�����Ȃ��AReadOnly
    Application.EnableEvents = True
    If objBook Is Nothing Then
        Call MsgBox(vFileName & vbCrLf & " ���J���܂���ł����B")
        Exit Sub
    End If
    On Error GoTo 0
    
    '�A�h�C���̎��͐ݒ�v
    objBook.IsAddin = True

'    '�t�H���_���ƃt�@�C�������擾
'    Dim strFolder As String '�t�H���_��
'    Dim strFile   As String '�t�@�C����(�g���q����)
'    With CreateObject("Scripting.FileSystemObject")
'        strFolder = .GetParentFolderName(vFileName)
'        strFile = .GetBaseName(vFileName)
'    End With
    
    '�g���q��.xlam�ɒu��
    Dim strAddinFile As String
    With CreateObject("Scripting.FileSystemObject")
        strAddinFile = Replace(vFileName, "." & .GetExtensionName(vFileName), ".xlam")
    End With
    
    Dim strPassword As String
    strPassword = InputBox("�ݒ肷��p�X���[�h����͂��Ă�������")
    
    Call objBook.SaveAs(strAddinFile, xlOpenXMLAddIn, strPassword)
    Call objBook.Close(False) '�ۑ������ɕ���
    
    Call MsgBox("�l���̍폜�@��" & vbLf & "�o�[�W�����ԍ��̕ύX��Y��Ȃ��ł�������")
End Sub
