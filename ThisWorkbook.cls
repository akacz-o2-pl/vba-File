VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public errState As Boolean
Public errMsg As String
Private Const modName As String = "Test."
Dim x

Sub main()
Const procName As String = "Main()"
Dim msg As String
Dim dirPath As String
Dim fileName As String
Dim fullFileName As String
Dim dirName As String
Dim fullDirName As String

On Error GoTo ErrorHandler
fileName = "A2018-03-06.txt"
dirPath = "C:\Documents and Settings\ak\Moje dokumenty\Kopie\MZB"
fullFileName = dirPath & "\" & fileName

' 1  File.ShowFolderList (dirPath)
' 2  MsgBox File.FileExists(fullFileName)
' 3  MsgBox Dir(fullFileName)
'4   MsgBox file.DriveExists("t:")
'   Dim text As Object
' 5  Set text = GetTextStream(fullFileName)
'    MsgBox text.readAll
'    Dim f As Object
' 6  Set f = GetFile(fullFileName)
'   MsgBox f.OpenAsTextStream().readAll
' 7    file.FileInfo fullFileName

'    fileName = "__2018-03-06.txt"
'    fullFileName = dirPath & "\" & fileName
'8   x = DeleteFile(fullFileName)

'  ? Czy funkcja CreateDir nie powinna zwróciæ nowo utworzonego katalogu??
'    dirPath = "C:\Documents and Settings\ak\Moje dokumenty"
'    dirName = "_Nowy_katalog_"
'    fullDirName = dirPath & "\" & dirName
'9   x = CreateDir(fullDirName)
'10  If x Then DeleteDir (fullDirName)

11 MsgBox "Walidacja: " & DirPathValidate(dirPath)

Done:
Exit Sub

ErrorHandler:
    If Not errState Then
        errState = True
        Err.Source = modName & procName & " linia " & Erl
    Else
        Err.Source = Err.Source & vbNewLine & _
                     modName & procName & " linia " & Erl
    End If
    msg = "B³¹d nr: " & Err.Number & vbNewLine & _
            Err.Description & vbNewLine & _
            vbNewLine & Err.Source
    LogError msg
    MsgBox msg
End Sub

' Rejestrowanie b³êdów VBA
Sub LogError(msg As String)
    Dim fileName As String, fileNo As Integer
    fileNo = FreeFile 'Get first free file number
    fileName = ThisWorkbook.Path & "\error_log.txt"
    Open fileName For Append As #fileNo
    Print #fileNo, Now & ":" & msg
    Close #fileNo
End Sub

