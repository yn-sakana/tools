Attribute VB_Name = "ExcelOperationLog"
' =============================================================================
' ExcelOperationLog.bas - Excel VBA
' Excelの操作ログを取得・記録する
'
' 使い方:
'   1. 監視したいブックのVBAエディタにこのモジュールをインポート
'   2. ThisWorkbook に以下のイベントハンドラを追加:
'
'   Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'       ExcelOperationLog.LogChange Sh, Target
'   End Sub
'
'   Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'       ExcelOperationLog.LogSave
'   End Sub
'
'   Private Sub Workbook_Open()
'       ExcelOperationLog.LogOpen
'   End Sub
'
'   Private Sub Workbook_NewSheet(ByVal Sh As Object)
'       ExcelOperationLog.LogNewSheet Sh
'   End Sub
' =============================================================================

Option Explicit

' --- 設定 ---
Private Const LOG_FILE As String = "C:\workspace\dev\tools\Data\logs.csv"

' =============================================================================
' ログ書き込み共通処理
' =============================================================================
Private Sub WriteLog(action As String, target As String, detail As String)
    Dim f As Integer
    Dim needHeader As Boolean

    ' ファイルが存在しなければヘッダーを書く
    needHeader = (Dir(LOG_FILE) = "")

    f = FreeFile
    Open LOG_FILE For Append As #f

    If needHeader Then
        Print #f, "timestamp,user,action,target,detail"
    End If

    Dim timestamp As String
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    Dim userName As String
    userName = Environ("USERNAME")

    Dim bookName As String
    bookName = ThisWorkbook.Name

    ' CSVエスケープ
    detail = Replace(detail, """", """""")
    If InStr(detail, ",") > 0 Or InStr(detail, """") > 0 Then
        detail = """" & detail & """"
    End If

    Print #f, timestamp & "," & userName & "," & action & "," & bookName & "," & detail
    Close #f
End Sub

' =============================================================================
' セル変更ログ
' =============================================================================
Public Sub LogChange(sh As Object, target As Range)
    Dim addr As String
    addr = sh.Name & "!" & target.Address(False, False)

    Dim detail As String
    If target.Cells.Count = 1 Then
        detail = addr & " を " & CStr(target.Value) & " に変更"
    Else
        detail = addr & " の範囲 (" & target.Cells.Count & " セル) を変更"
    End If

    WriteLog "セル編集", ThisWorkbook.Name, detail
End Sub

' =============================================================================
' 保存ログ
' =============================================================================
Public Sub LogSave()
    WriteLog "ファイル保存", ThisWorkbook.Name, "上書き保存"
End Sub

' =============================================================================
' ファイルオープンログ
' =============================================================================
Public Sub LogOpen()
    Dim detail As String
    If ThisWorkbook.ReadOnly Then
        detail = "読み取り専用で開く"
    Else
        detail = "編集モードで開く"
    End If
    WriteLog "ファイル開く", ThisWorkbook.Name, detail
End Sub

' =============================================================================
' シート追加ログ
' =============================================================================
Public Sub LogNewSheet(sh As Object)
    WriteLog "シート追加", ThisWorkbook.Name, sh.Name & " を追加"
End Sub
