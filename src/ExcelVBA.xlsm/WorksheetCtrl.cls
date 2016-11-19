VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' @brief    Check worksheet
'
' @param    sName       Target worksheet name
'
' @return   Boolean     True: Exist / False: Nothing
'
Public Function Check(ByVal sTarget As String) As Boolean

    Dim i, cnt As Integer
    
    cnt = Sheets.Count
    CheckSheet = False
    
    For i = 1 To cnt
        If Sheets(i).Name = sTarget Then
            CheckSheet = True
            Exit For
        End If
    Next

End Function

''
' @brief    Create worksheet
'
' @param    [sNew)      Create wroksheet name
' @param    [sInsert]   Insertion position of worksheet name
' @param    [sActive]   Active worksheet name after create
'
' @return   none
'
' @note     Checkメソッドと合わせて使うと便利
'
Public Sub Create(Optional ByVal sNew As String, Optional ByVal sInsert As String, Optional ByVal sActive As String)

    ' Create worksheet
    If sInsert <> "" Then
        Worksheets.Add after:=Worksheets(sInsert)
    Else
        Worksheets.Add
    End If
    
    ' Rename worksheet
    If sNew <> "" Then
        ActiveSheet.Name = sNew
    End If
    
    ' Active worksheet
    If sActive <> "" Then
        Worksheets(sActive).Active
    End If

End Sub

''
' @brief    Import CSV
'
' @param    sCsv        CSV file name
' @param    [sCell]     Target cell (Start cell)
'
' @return   none
'
' @note     CSVファイルを自動インポートすると16進数が対数表現になるのを回避する用途。
'
Public Sub QueryCSV(ByVal sCsv As String, Optional ByVal sCell As String)

    Dim cell As String
    
    If sCell <> "" Then
        cell = sCell
    Else
        cell = "A1"
    End If
    
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & sCsv, Destination:=ActiveSheet.Range(cell))
        .FieldNames = True
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshStyle = xlInsertDeleteCells
        .AdjustColumnWidth = True
        .TextFilePlatform = 932
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
        .Parent.Names(.Name).Delete
        .Delete
    End With

End Sub

