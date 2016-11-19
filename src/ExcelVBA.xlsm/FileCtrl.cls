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
' @brief    Specific range substitution of external files
'
' @param    sPath       Operating target file path
' @param    sKeyS       String of "StartKEY"
' @param    sKeyE       String of "EndKEY"
' @param    sStr        String of replace
' @param    [bBackup]   Backup (True:Execute / False:Nothing)
'
' @return   none
'
' @note     外部ファイルの"開始キー"から"終了キー"の範囲を置換する
'
Public Sub Replace4Area(ByVal sTarget As String, ByVal sKeyS As String, ByVal sKeyE As String, ByVal sStr As String, Optional ByVal bBackup As Boolean)

    Const FOR_READING = 1
    Const FOR_WRITING = 2
    Const FOR_APPENDING = 8
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Open file
    Dim SRC As Object
    Set SRC = FSO.OpenTextFile(sTarget, FOR_READING)
    
    Dim dstPath As String
    dstPath = sTarget + ".tmp"
    
    Dim DST As Object
    Set DST = FSO.OpenTextFile(dstPath, FOR_WRITING, True)
    
    ' Copy file
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
    
    ' Close file
    Set SRC = Nothing
    Set DST = Nothing
    
    ' Kill(Rename) file
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
