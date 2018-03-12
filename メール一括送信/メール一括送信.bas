Attribute VB_Name = "Module1"
Public strAdd As String            '添付ファイル


Sub Mail()

Dim strForm As String           '送信メールアドレス
Dim strTo As String             '送信先メールアドレス
Dim strSubject As String        'メールタイトル名
Dim strTextBody As String       'メール本文

Dim intSend As Integer          '送信方法（1：ローカルSMTP、2：SMTPポート、3：OLE DB)
Dim strServer As String         'SMTPサーバー名
Dim intPort As Integer          'SMTPポート番号
Dim boSSL As Boolean            'SSL通信
Dim intCate As Integer          'SMTP認証（1：Basic認証、2：NTLM認証）
Dim strUser As String           'ユーザー名
Dim strPASS As String           'パスワード
Dim intTimeout As Integer       '接続タイムアウト秒数

Dim strSetteiSheet As String    '設定シート名
Dim strSoushinSheet As String   '送信シート名

Dim intKensu As Integer         '送信数
Dim intErr As Integer           'エラー数
'-------------------------------------------
'       初期値
'-------------------------------------------
intSend = 2                         '送信方法(2:SMTPポート）
intTimeout = 60                     '接続タイムアウト秒数

strSetteiSheet = "送信"             '設定シート名
strSoushinSheet = "送信者一覧"      '送信シート名

strConfigurationField = "http://schemas.microsoft.com/cdo/configuration/"

intErr = 0                          'エラー数初期化

'-------------------------------------------
'       データ取得
'-------------------------------------------
'送信メールアドレス
strForm = worksheets(strSetteiSheet).Cells(2, 2).Text & "<" & worksheets(strSetteiSheet).Cells(1, 2).Text & ">"
'送信サーバー名
strServer = worksheets(strSetteiSheet).Cells(3, 2).Text
'SMTPポート
intPort = worksheets(strSetteiSheet).Cells(4, 2).Text
'SSL通信
boSSL = worksheets(strSetteiSheet).ckSSL.Value
'SMTP認証
If worksheets(strSetteiSheet).opSMTP1.Value = True Then
    intCate = 1
ElseIf worksheets(strSetteiSheet).opSMTP2.Value = True Then
    intCate = 2
Else
    'どちらも入っていない場合仮に1：Basic認証とする
    intCate = 1
End If
'送信ユーザー名
strUser = worksheets(strSetteiSheet).Cells(7, 2).Text
'送信パスワード
strPASS = worksheets(strSetteiSheet).Cells(8, 2).Text
'送信タイトル
strSubject = worksheets(strSetteiSheet).Cells(10, 2).Text
'送信本文
strTextBody = worksheets(strSetteiSheet).Cells(11, 2).Text

'送信数（行数で取得の為マイナス１）
intKensu = worksheets(strSoushinSheet).Cells(2, 3).End(xlDown).Row - 1

'確認画面
If MsgBox(intKensu & "件" & vbCrLf & strAdd & vbCrLf & "メール送信しますか？", vbYesNo) = vbYes Then

    'イエスの時のみ発動
    For i = 1 To intKensu
    
        '送信先アドレス
        strTo = worksheets(strSoushinSheet).Cells(i + 1, 3).Text
        
        '送信先アドレス確認
        If strTo = "" Then
            '空欄時エラー数増減
            intErr = intErr + 1
        Else
    
            'メール送信　送信時　名前＋様を追加
            Call MailAdd(strForm, strTo, strSubject, worksheets(strSoushinSheet).Cells(i + 1, 2).Text & "様" & vbCrLf & strTextBody, _
                        strAdd, intSend, strServer, intPort, boSSL, intCate, strUser, strPASS, intTimeout, strConfigurationField)
        End If
        
    Next
    
    'エラー数があるかどうか
    If intErr > 0 Then
        MsgBox (intKensu & "件中 " & intErr & "件" & vbCrLf & "送信できませんでした。")
    Else
        MsgBox ("送信完了しました")
    End If
End If

End Sub

'*****************************************************
'   添付
'*****************************************************
Sub cmAddClick()
    '添付
    strAdd = FileName
    
End Sub


'******************************************************
'   ファイル名取得
'   filename 引き渡し
'******************************************************

Function FileName() As String
      
    '=====================
    '   ファイル指定
    '=====================
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "ファイルの選択"
        'ファイルの種類を設定
        .Filters.Clear
        .Filters.Add "すべてのファイル", "*.*"
        '複数ファイル選択を許可しない
        .AllowMultiSelect = False
          
        'ダイアログを表示
        If .Show = -1 Then
            'ファイルが選択されたとき
            'そのフルバスを返り値に設定
            FileName = Trim(.SelectedItems.Item(1))
        Else
            'ファイルが選択されなければ長さゼロの文字列を返す
            FileName = ""
        End If
    End With
           
End Function

'******************************************************
'   VBScriptのCDO.Message
'******************************************************

Sub MailAdd(strForm As String, strTo As String, strSubject As String, strTextBody As String, strAdd As String, intSend As Integer, strServer As String, _
            intPort As Integer, boSSL As Boolean, intCate As Integer, strUser As String, strPASS As String, intTimeout As Integer, ByRef strConfigurationField As Variant)
    
    Dim strBody As String       '仮本文
    
    '-------------------------
    ' 本文の改行コードの確認
    '-------------------------
    ' Lfのみの場合Cr+Lfに変換
    strBody = Replace(strTextBody, vbLf, vbCrLf)
    ' 上記で元がCr+Lfの場合Cr+Cr+LfになるのでCr+Lfに戻す
    strTextBody = Replace(strBody, vbCr & vbCrLf, vbCrLf)
    
    '----------------------------------
    '   送信設定
    '----------------------------------
    Set objMail = CreateObject("CDO.Message")
    
    objMail.From = strForm
    objMail.To = strTo
    objMail.Subject = strSubject
    objMail.TextBody = strTextBody
    '添付確認
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
