VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' @brief    外部ファイルの特定範囲の文字列置換
'
' @param    sPath       操作対象となるファイルパス
' @param    sKeyS       開始キーとなる文字列
' @param    sKeyE       終了キーとなる文字列
' @param    sStr        置換する文字列
' @param    bBackup     バックアップ (True:実行 / False:実行しない)
'
' @return   none
'
' @note
'
Sub replace4area(sTarget As String, sKeyS As String, sKeyE As String, sStr As String, bBackup As Boolean)

    Const FOR_READING = 1
    Const FOR_WRITING = 2
    Const FOR_APPENDING = 8
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' File Open
    Dim SRC As Object
    Set SRC = FSO.OpenTextFile(sTarget, FOR_READING)
    
    Dim dstPath As String
    dstPath = sTarget + ".tmp"
    
    Dim DST As Object
    Set DST = FSO.OpenTextFile(dstPath, FOR_WRITING, True)
    
    ' File Copy
    Dim ReadLine As String
    Dim skip As Boolean
    
    skip = False
    
    Do Until SRC.atEndOfStream = True
    
        ReadLine = SRC.ReadLine
        
        Select Case ReadLine
            Case sKeyS
                skip = True
                DST.WriteLine (ReadLine)
                DST.WriteLine (sStr)
            Case sKeyE
                skip = False
                DST.WriteLine (ReadLine)
            Case Else
                If skip = False Then
                    DST.WriteLine (ReadLine)
                End If
        End Select
        
    Loop
    
    ' File Close
    Set SRC = Nothing
    Set DST = Nothing
    
    ' File Kill & ReName
    If bBackup = True Then
        Dim pathBackup As String
        pathBackup = sTarget & ".bak"
        If Dir(pathBackup) <> "" Then
            Kill pathBackup
        End If
        
        Name sTarget As pathBackup
        
    Else
        Kill sTarget
    
    End If
    
    Name dstPath As sTarget

End Sub
