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
' @brief    �O���t�@�C���̓���͈͂̕�����u��
'
' @param    sPath       ����ΏۂƂȂ�t�@�C���p�X
' @param    sKeyS       �J�n�L�[�ƂȂ镶����
' @param    sKeyE       �I���L�[�ƂȂ镶����
' @param    sStr        �u�����镶����
' @param    bBackup     �o�b�N�A�b�v (True:���s / False:���s���Ȃ�)
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