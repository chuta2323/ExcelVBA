VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellNameFommater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ALPHABET_MAX = 26
Private Const CODE_A = 65
    

''
' @brief    Convert column from Integer to Alphabet
'
' @param    iCol        Integer in column
'
' @return   String      Alphabet in Column (MAX: "IV")
'
Public Function ConvertColumn2Alphabet(ByVal iCol As Integer) As String

    Dim alphabet As Integer: alphabet = Int((iCol - 1) / (CODE_A - 1))
    Dim remainder As Integer: remainder = iCol - (alphabet * ALPHABET_MAX)
    
    If alphabet > 0 Then
        ConvertColumn2Alphabet = Chr(alphabet + (CODE_A - 1))
    End If
    
    If remainder > 0 Then
        ConvertColumn2Alphabet = ConvertColumn2Alphabet + Chr(remainder + (CODE_A - 1))
    End If

End Function

''
' @brief    Convert column from Alphabet to Integer
'
' @param    sCol        Alphabet in column (MAX: "IV")
'
' @return   Integer     Integer in Column
'
Public Function ConvertColumn2Integer(ByVal sCol As String) As Integer

    Dim number As Integer: number = AscW(Left(sCol, 1)) - (CODE_A - 1)
    Dim length As Integer: length = Len(sCol)
    
    If length = 1 Then
        ConvertColumn2Integer = number
    Else
        Dim figure As Integer: figure = number * ALPHABET_MAX
        number = AscW(Right(sCol, 1)) - (CODE_A - 1)
        ConvertColumn2Integer = figure + number
    End If

End Function

''
' @brief    Extract row from cell
'
' @param    sCellName   Cell name
'
' @return   Integer     Row
'
Public Function ExtractRowFromCell(ByVal sCellName As String) As Integer

    Dim cPos As Integer: cPos = 1
    Dim sepPos As Integer
    Dim checkChr As String
    Dim ascii As Integer
    
    Do While cPos <= Len(sCellName)
        checkChr = Mid(sCellName, cPos, 1)
        ascii = Asc(checkChr)
        If 48 <= ascii And ascii <= 57 Then
            Dim tempStr As String: tempStr = Mid(sCellName, cPos, Len(sCellName))
            ExtractRowFromCell = tempStr
            Exit Do
        End If
        cPos = cPos + 1
    Loop

End Function

''
' @brief    Extract column from cell
'
' @param    sCellName   Cell name
'
' @return   String      Alphabet in column
'
Public Function ExtractColumnFromCell(ByVal sCellName As String) As String

    Dim cPos As Integer: cPos = 1
    Dim sepPos As Integer
    Dim checkChr As String
    Dim ascii As Integer
    
    Do While cPos <= Len(sCellName)
        checkChr = Mid(sCellName, cPos, 1)
        ascii = Asc(checkChr)
        If 48 <= ascii And ascii <= 57 Then
            Dim tempStr As String: tempStr = Mid(sCellName, 1, (cPos - 1))
            ExtractColumnFromCell = tempStr
            Exit Do
        End If
        cPos = cPos + 1
    Loop

End Function

