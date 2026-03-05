Attribute VB_Name = "ExcelExporter"
' =============================================================================
' ExcelExporter.bas - Excel VBA
' ブック内のテーブル（ListObject）をCSV/JSONとして自動エクスポートする
'
' 使い方:
'   1. Excel VBAエディタ（Alt+F11）にこのモジュールをインポート
'   2. ExportAllTables を手動実行、または ThisWorkbook に BeforeSave イベントを設定
'
' ThisWorkbook に以下を追加すると保存時に自動エクスポートされる:
'   Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'       ExcelExporter.ExportAllTables
'   End Sub
' =============================================================================

Option Explicit

' --- 設定 ---
Private Const EXPORT_DIR As String = "C:\workspace\dev\tools\Data"

' =============================================================================
' 全テーブルをエクスポート
' =============================================================================
Public Sub ExportAllTables()
    Dim ws As Worksheet
    Dim tbl As ListObject

    ' エクスポート先フォルダ作成
    CreateFolderIfNotExists EXPORT_DIR

    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            ' CSVエクスポート
            ExportTableToCsv tbl, EXPORT_DIR & "\" & tbl.Name & ".csv"
            ' JSONエクスポート
            ExportTableToJson tbl, EXPORT_DIR & "\" & tbl.Name & ".json"
        Next tbl
    Next ws
End Sub

' =============================================================================
' テーブルをCSVにエクスポート
' =============================================================================
Private Sub ExportTableToCsv(tbl As ListObject, filePath As String)
    Dim f As Integer
    f = FreeFile
    Open filePath For Output As #f

    ' ヘッダー行
    Dim headers As String
    Dim col As ListColumn
    Dim first As Boolean
    first = True
    For Each col In tbl.ListColumns
        If Not first Then headers = headers & ","
        headers = headers & EscapeCsvField(col.Name)
        first = False
    Next col
    Print #f, headers

    ' データ行
    Dim row As ListRow
    For Each row In tbl.ListRows
        Dim line As String
        line = ""
        first = True
        Dim c As Long
        For c = 1 To tbl.ListColumns.Count
            If Not first Then line = line & ","
            Dim cellValue As String
            cellValue = CStr(row.Range(1, c).Value)
            line = line & EscapeCsvField(cellValue)
            first = False
        Next c
        Print #f, line
    Next row

    Close #f
End Sub

' =============================================================================
' テーブルをJSONにエクスポート
' =============================================================================
Private Sub ExportTableToJson(tbl As ListObject, filePath As String)
    Dim f As Integer
    f = FreeFile
    Open filePath For Output As #f

    Print #f, "["

    Dim row As ListRow
    Dim rowIdx As Long
    rowIdx = 0

    For Each row In tbl.ListRows
        rowIdx = rowIdx + 1

        Dim line As String
        line = "  {"

        Dim c As Long
        For c = 1 To tbl.ListColumns.Count
            Dim colName As String
            colName = tbl.ListColumns(c).Name

            Dim cellValue As String
            cellValue = CStr(row.Range(1, c).Value)

            ' JSONエスケープ
            cellValue = Replace(cellValue, "\", "\\")
            cellValue = Replace(cellValue, """", "\""")
            cellValue = Replace(cellValue, vbCr, "")
            cellValue = Replace(cellValue, vbLf, "\n")

            If c > 1 Then line = line & ", "
            line = line & """" & colName & """: """ & cellValue & """"
        Next c

        line = line & "}"
        If rowIdx < tbl.ListRows.Count Then line = line & ","

        Print #f, line
    Next row

    Print #f, "]"
    Close #f
End Sub

' =============================================================================
' CSV用フィールドエスケープ
' =============================================================================
Private Function EscapeCsvField(value As String) As String
    If InStr(value, ",") > 0 Or InStr(value, """") > 0 Or InStr(value, vbLf) > 0 Then
        value = Replace(value, """", """""")
        EscapeCsvField = """" & value & """"
    Else
        EscapeCsvField = value
    End If
End Function

' =============================================================================
' フォルダが無ければ作成
' =============================================================================
Private Sub CreateFolderIfNotExists(folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    Set fso = Nothing
End Sub
