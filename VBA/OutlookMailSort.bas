Attribute VB_Name = "OutlookMailSort"
' =============================================================================
' OutlookMailSort.bas - Outlook VBA
' 受信メールを連絡先別にフォルダ保存（添付ファイル含む）
'
' 使い方:
'   1. Outlook VBAエディタ（Alt+F11）にこのモジュールをインポート
'   2. 連絡先Excelファイルのパスを CONTACTS_FILE 定数に設定
'   3. SortMailsByContact を実行
' =============================================================================

Option Explicit

' --- 設定 ---
Private Const CONTACTS_FILE As String = "C:\workspace\dev\tools\Data\contacts.json"
Private Const DEFAULT_SAVE_FOLDER As String = "C:\workspace\dev\tools\Tests\MailArchive\未分類"

' --- 連絡先を格納する型 ---
Private Type ContactInfo
    Name As String
    Email As String
    Company As String
    SaveFolder As String
End Type

' =============================================================================
' メイン処理: 受信トレイのメールを連絡先別に保存
' =============================================================================
Public Sub SortMailsByContact()
    Dim contacts() As ContactInfo
    contacts = LoadContactsFromJson(CONTACTS_FILE)

    If UBound(contacts) < 0 Then
        MsgBox "連絡先が読み込めませんでした", vbExclamation
        Exit Sub
    End If

    Dim olApp As Outlook.Application
    Set olApp = Outlook.Application

    Dim olNs As Outlook.NameSpace
    Set olNs = olApp.GetNamespace("MAPI")

    Dim inbox As Outlook.MAPIFolder
    Set inbox = olNs.GetDefaultFolder(olFolderInbox)

    Dim mailItem As Object
    Dim savedCount As Long
    savedCount = 0

    Dim i As Long
    For i = inbox.Items.Count To 1 Step -1
        Set mailItem = inbox.Items(i)

        If TypeOf mailItem Is Outlook.mailItem Then
            Dim mail As Outlook.mailItem
            Set mail = mailItem

            Dim senderEmail As String
            senderEmail = GetSenderEmail(mail)

            Dim saveFolder As String
            saveFolder = FindSaveFolder(senderEmail, contacts)

            ' フォルダ作成
            CreateFolderIfNotExists saveFolder

            ' メール本文を .txt で保存
            Dim safeSubject As String
            safeSubject = SanitizeFileName(mail.Subject)

            Dim dateStr As String
            dateStr = Format(mail.ReceivedTime, "yyyymmdd_hhnnss")

            Dim baseName As String
            baseName = dateStr & "_" & safeSubject

            Dim txtPath As String
            txtPath = saveFolder & "\" & baseName & ".txt"

            Dim f As Integer
            f = FreeFile
            Open txtPath For Output As #f
            Print #f, "From: " & mail.SenderName & " <" & senderEmail & ">"
            Print #f, "To: " & mail.To
            Print #f, "Date: " & Format(mail.ReceivedTime, "yyyy/mm/dd hh:nn:ss")
            Print #f, "Subject: " & mail.Subject
            Print #f, String(60, "-")
            Print #f, mail.Body
            Close #f

            ' 添付ファイル保存
            If mail.Attachments.Count > 0 Then
                Dim attachDir As String
                attachDir = saveFolder & "\" & baseName & "_添付"
                CreateFolderIfNotExists attachDir

                Dim att As Outlook.Attachment
                Dim j As Long
                For j = 1 To mail.Attachments.Count
                    Set att = mail.Attachments(j)
                    ' 埋め込み画像等はスキップ
                    If att.Type <> olEmbeddeditem Then
                        att.SaveAsFile attachDir & "\" & att.FileName
                    End If
                Next j
            End If

            savedCount = savedCount + 1
        End If
    Next i

    MsgBox savedCount & " 件のメールを保存しました", vbInformation
End Sub

' =============================================================================
' 連絡先JSONファイルを読み込む（簡易JSONパーサー）
' =============================================================================
Private Function LoadContactsFromJson(filePath As String) As ContactInfo()
    Dim contacts() As ContactInfo
    ReDim contacts(-1 To -1)

    If Dir(filePath) = "" Then
        LoadContactsFromJson = contacts
        Exit Function
    End If

    Dim f As Integer
    f = FreeFile
    Dim jsonText As String
    Dim line As String

    Open filePath For Input As #f
    Do Until EOF(f)
        Line Input #f, line
        jsonText = jsonText & line
    Loop
    Close #f

    ' 簡易パース: 各オブジェクトを抽出
    Dim idx As Long
    idx = 0
    Dim pos As Long
    pos = 1

    Do
        Dim objStart As Long
        objStart = InStr(pos, jsonText, "{")
        If objStart = 0 Then Exit Do

        Dim objEnd As Long
        objEnd = InStr(objStart, jsonText, "}")
        If objEnd = 0 Then Exit Do

        Dim objText As String
        objText = Mid(jsonText, objStart, objEnd - objStart + 1)

        ReDim Preserve contacts(0 To idx)
        contacts(idx).Name = ExtractJsonValue(objText, "name")
        contacts(idx).Email = ExtractJsonValue(objText, "email")
        contacts(idx).Company = ExtractJsonValue(objText, "company")
        contacts(idx).SaveFolder = ExtractJsonValue(objText, "saveFolder")

        idx = idx + 1
        pos = objEnd + 1
    Loop

    LoadContactsFromJson = contacts
End Function

' =============================================================================
' JSON文字列から指定キーの値を抽出（簡易）
' =============================================================================
Private Function ExtractJsonValue(jsonObj As String, key As String) As String
    Dim searchKey As String
    searchKey = """" & key & """"

    Dim keyPos As Long
    keyPos = InStr(1, jsonObj, searchKey)
    If keyPos = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If

    Dim colonPos As Long
    colonPos = InStr(keyPos, jsonObj, ":")
    If colonPos = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If

    Dim valStart As Long
    valStart = InStr(colonPos, jsonObj, """")
    If valStart = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If
    valStart = valStart + 1

    Dim valEnd As Long
    valEnd = InStr(valStart, jsonObj, """")

    Dim rawValue As String
    rawValue = Mid(jsonObj, valStart, valEnd - valStart)

    ' エスケープされたバックスラッシュを復元
    rawValue = Replace(rawValue, "\\", "\")

    ExtractJsonValue = rawValue
End Function

' =============================================================================
' 送信者メールアドレスを取得
' =============================================================================
Private Function GetSenderEmail(mail As Outlook.mailItem) As String
    If mail.SenderEmailType = "EX" Then
        ' Exchange の場合は SMTP アドレスに変換
        Dim sender As Outlook.AddressEntry
        Set sender = mail.sender
        If Not sender Is Nothing Then
            Dim exUser As Outlook.ExchangeUser
            Set exUser = sender.GetExchangeUser()
            If Not exUser Is Nothing Then
                GetSenderEmail = LCase(exUser.PrimarySmtpAddress)
                Exit Function
            End If
        End If
        GetSenderEmail = LCase(mail.SenderEmailAddress)
    Else
        GetSenderEmail = LCase(mail.SenderEmailAddress)
    End If
End Function

' =============================================================================
' メールアドレスから保存先フォルダを検索
' =============================================================================
Private Function FindSaveFolder(email As String, contacts() As ContactInfo) As String
    Dim i As Long
    For i = LBound(contacts) To UBound(contacts)
        If LCase(contacts(i).Email) = LCase(email) Then
            FindSaveFolder = contacts(i).SaveFolder
            Exit Function
        End If
    Next i
    FindSaveFolder = DEFAULT_SAVE_FOLDER
End Function

' =============================================================================
' ファイル名に使えない文字を除去
' =============================================================================
Private Function SanitizeFileName(name As String) As String
    Dim result As String
    result = name
    Dim invalid As Variant
    For Each invalid In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        result = Replace(result, invalid, "_")
    Next invalid
    If Len(result) > 100 Then result = Left(result, 100)
    SanitizeFileName = result
End Function

' =============================================================================
' フォルダが無ければ再帰的に作成
' =============================================================================
Private Sub CreateFolderIfNotExists(folderPath As String)
    If Dir(folderPath, vbDirectory) = "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(folderPath) Then
            CreateFolderIfNotExists fso.GetParentFolderName(folderPath)
            MkDir folderPath
        End If
        Set fso = Nothing
    End If
End Sub
