Attribute VB_Name = "Module1"
Public strAdd As String            '�Y�t�t�@�C��


Sub Mail()

Dim strForm As String           '���M���[���A�h���X
Dim strTo As String             '���M�惁�[���A�h���X
Dim strSubject As String        '���[���^�C�g����
Dim strTextBody As String       '���[���{��

Dim intSend As Integer          '���M���@�i1�F���[�J��SMTP�A2�FSMTP�|�[�g�A3�FOLE DB)
Dim strServer As String         'SMTP�T�[�o�[��
Dim intPort As Integer          'SMTP�|�[�g�ԍ�
Dim boSSL As Boolean            'SSL�ʐM
Dim intCate As Integer          'SMTP�F�؁i1�FBasic�F�؁A2�FNTLM�F�؁j
Dim strUser As String           '���[�U�[��
Dim strPASS As String           '�p�X���[�h
Dim intTimeout As Integer       '�ڑ��^�C���A�E�g�b��

Dim strSetteiSheet As String    '�ݒ�V�[�g��
Dim strSoushinSheet As String   '���M�V�[�g��

Dim intKensu As Integer         '���M��
Dim intErr As Integer           '�G���[��
'-------------------------------------------
'       �����l
'-------------------------------------------
intSend = 2                         '���M���@(2:SMTP�|�[�g�j
intTimeout = 60                     '�ڑ��^�C���A�E�g�b��

strSetteiSheet = "���M"             '�ݒ�V�[�g��
strSoushinSheet = "���M�҈ꗗ"      '���M�V�[�g��

strConfigurationField = "http://schemas.microsoft.com/cdo/configuration/"

intErr = 0                          '�G���[��������

'-------------------------------------------
'       �f�[�^�擾
'-------------------------------------------
'���M���[���A�h���X
strForm = worksheets(strSetteiSheet).Cells(2, 2).Text & "<" & worksheets(strSetteiSheet).Cells(1, 2).Text & ">"
'���M�T�[�o�[��
strServer = worksheets(strSetteiSheet).Cells(3, 2).Text
'SMTP�|�[�g
intPort = worksheets(strSetteiSheet).Cells(4, 2).Text
'SSL�ʐM
boSSL = worksheets(strSetteiSheet).ckSSL.Value
'SMTP�F��
If worksheets(strSetteiSheet).opSMTP1.Value = True Then
    intCate = 1
ElseIf worksheets(strSetteiSheet).opSMTP2.Value = True Then
    intCate = 2
Else
    '�ǂ���������Ă��Ȃ��ꍇ����1�FBasic�F�؂Ƃ���
    intCate = 1
End If
'���M���[�U�[��
strUser = worksheets(strSetteiSheet).Cells(7, 2).Text
'���M�p�X���[�h
strPASS = worksheets(strSetteiSheet).Cells(8, 2).Text
'���M�^�C�g��
strSubject = worksheets(strSetteiSheet).Cells(10, 2).Text
'���M�{��
strTextBody = worksheets(strSetteiSheet).Cells(11, 2).Text

'���M���i�s���Ŏ擾�̈׃}�C�i�X�P�j
intKensu = worksheets(strSoushinSheet).Cells(2, 3).End(xlDown).Row - 1

'�m�F���
If MsgBox(intKensu & "��" & vbCrLf & strAdd & vbCrLf & "���[�����M���܂����H", vbYesNo) = vbYes Then

    '�C�G�X�̎��̂ݔ���
    For i = 1 To intKensu
    
        '���M��A�h���X
        strTo = worksheets(strSoushinSheet).Cells(i + 1, 3).Text
        
        '���M��A�h���X�m�F
        If strTo = "" Then
            '�󗓎��G���[������
            intErr = intErr + 1
        Else
    
            '���[�����M�@���M���@���O�{�l��ǉ�
            Call MailAdd(strForm, strTo, strSubject, worksheets(strSoushinSheet).Cells(i + 1, 2).Text & "�l" & vbCrLf & strTextBody, _
                        strAdd, intSend, strServer, intPort, boSSL, intCate, strUser, strPASS, intTimeout, strConfigurationField)
        End If
        
    Next
    
    '�G���[�������邩�ǂ���
    If intErr > 0 Then
        MsgBox (intKensu & "���� " & intErr & "��" & vbCrLf & "���M�ł��܂���ł����B")
    Else
        MsgBox ("���M�������܂���")
    End If
End If

End Sub

'*****************************************************
'   �Y�t
'*****************************************************
Sub cmAddClick()
    '�Y�t
    strAdd = FileName
    
End Sub


'******************************************************
'   �t�@�C�����擾
'   filename �����n��
'******************************************************

Function FileName() As String
      
    '=====================
    '   �t�@�C���w��
    '=====================
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "�t�@�C���̑I��"
        '�t�@�C���̎�ނ�ݒ�
        .Filters.Clear
        .Filters.Add "���ׂẴt�@�C��", "*.*"
        '�����t�@�C���I���������Ȃ�
        .AllowMultiSelect = False
          
        '�_�C�A���O��\��
        If .Show = -1 Then
            '�t�@�C�����I�����ꂽ�Ƃ�
            '���̃t���o�X��Ԃ�l�ɐݒ�
            FileName = Trim(.SelectedItems.Item(1))
        Else
            '�t�@�C�����I������Ȃ���Β����[���̕������Ԃ�
            FileName = ""
        End If
    End With
           
End Function

'******************************************************
'   VBScript��CDO.Message
'******************************************************

Sub MailAdd(strForm As String, strTo As String, strSubject As String, strTextBody As String, strAdd As String, intSend As Integer, strServer As String, _
            intPort As Integer, boSSL As Boolean, intCate As Integer, strUser As String, strPASS As String, intTimeout As Integer, ByRef strConfigurationField As Variant)
    
    Dim strBody As String       '���{��
    
    '-------------------------
    ' �{���̉��s�R�[�h�̊m�F
    '-------------------------
    ' Lf�݂̂̏ꍇCr+Lf�ɕϊ�
    strBody = Replace(strTextBody, vbLf, vbCrLf)
    ' ��L�Ō���Cr+Lf�̏ꍇCr+Cr+Lf�ɂȂ�̂�Cr+Lf�ɖ߂�
    strTextBody = Replace(strBody, vbCr & vbCrLf, vbCrLf)
    
    '----------------------------------
    '   ���M�ݒ�
    '----------------------------------
    Set objMail = CreateObject("CDO.Message")
    
    objMail.From = strForm
    objMail.To = strTo
    objMail.Subject = strSubject
    objMail.TextBody = strTextBody
    '�Y�t�m�F
    If strAdd <> "" Then
        objMail.AddAttachment strAdd
    End If

    With objMail.Configuration.Fields
        .Item(strConfigurationField & "sendusing") = intSend
        .Item(strConfigurationField & "smtpserver") = strServer
        .Item(strConfigurationField & "smtpserverport") = intPort
        .Item(strConfigurationField & "smtpusessl") = boSSL
        .Item(strConfigurationField & "smtpauthenticate") = intCate
        .Item(strConfigurationField & "sendusername") = strUser
        .Item(strConfigurationField & "sendpassword") = strPASS
        .Item(strConfigurationField & "smtpconnectiontimeout") = intTimeout
        .Update
    End With

objMail.Send

Set objMail = Nothing


End Sub
