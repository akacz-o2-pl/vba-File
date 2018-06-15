Attribute VB_Name = "File"
Option Explicit
'********************************************************************
'*                 Operacje na plikach i katalogach                 *
'********************************************************************
Public errState As Boolean
Public errMsg As String
Private Const modName As String = "File."
Dim x, xx
' https://www.w3schools.com/asp/asp_ref_filesystem.asp

' Sprawdza czy dysk istnieje
' @para String: driveName - nazwa dysku (c:)
' return Boolean: True | False
Function DriveExists(driveName) As Boolean
    Dim fs
    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    DriveExists = fs.DriveExists(driveName)
    Set fs = Nothing
End Function

' Sprawdza czy katalog istnieje
' @para String: dirPath - œcie¿ka dostêpu do katalogu.
' return Boolean: True | False
Function DirExists(dirPath) As Boolean
    On Error Resume Next
    DirExists = (GetAttr(dirPath) And vbDirectory) = vbDirectory
End Function

' Sprawdza czy plik istnieje
' @para String: fullFileName - pe³na œcie¿ka dostêpu do pliku wraz z jego nazw¹
' return Boolean: True | False
Function FileExists(ByVal fullFileName As String) As Boolean
FileExists = Dir(fullFileName) <> ""
End Function

' Wyœwietla msgBox z list¹ plików w podanym katalogu
' @para String: dirPath - œcie¿ka dostêpu do katalogu
' @return Object: MsgBox
Function ShowFolderList(dirPath)
Dim fs, folder, fc, file, result
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set folder = fs.GetFolder(dirPath)
    Set fc = folder.Files
    For Each file In fc
        result = result & file.Name
        result = result & vbCrLf
    Next
    MsgBox result
    'Debug.Print result
End Function

' Zwraca listê plików z podanego katalogu
' @para String: dirPath - œcie¿ka dostêpu do katalogu
' @return Object - lista plików
Public Function GetFileList(dirPath As String) As Object
Dim fs, folder, file
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set folder = fs.GetFolder(dirPath)
    Set GetFileList = folder.Files
End Function

' Wyœwietla informacje o pliku
' @param String fullFileName - pe³na œcie¿ka dostêpu do pliku
' return Object: Msgbox
Public Function FileInfo(fullFileName As String)
    Const procName As String = "FileInfo(fullFileName)"
    Dim fs, drive, folder, file
    Dim driveName As String, dirPath As String, fileName As String
    Dim fileBaseName, fileExt As String
    Dim msg As String

On Error GoTo ErrorHandler
    Set fs = CreateObject("Scripting.FileSystemObject")
    fileName = fs.GetFileName(fullFileName)
    fileBaseName = fs.GetBaseName(fullFileName)
    fileExt = fs.GetExtensionName(fullFileName)
'    dirPath = Left(fullFileName, InStrRev(fullFileName, "\"))
    dirPath = fs.GetParentFolderName(fullFileName)
    driveName = Left(fullFileName, 3)
 
1   Set drive = fs.GetDrive(driveName)
2   Set folder = fs.GetFolder(dirPath)
3   Set file = fs.GetFile(fullFileName)
    msg = "dysk: " & drive & vbNewLine & _
          "katalog: " & dirPath & vbNewLine & _
          "plik: " & fileName & vbNewLine & _
          "nazwa pliku: " & fileBaseName & vbNewLine & _
          "rozszerzenie: " & fileExt & vbNewLine
    MsgBox msg
    Set fs = Nothing
    Set drive = Nothing
    Set folder = Nothing
    Set file = Nothing
Done:
Exit Function

ErrorHandler:
    If Not errState Then
        errState = True
        Err.Source = modName & procName & " linia " & Erl
    Else
        Err.Source = Err.Source & vbNewLine & _
                     modName & procName & " linia " & Erl
    End If
    Select Case Err.Number
        Case 53
            Err.Description = "Nie znaleziono pliku: " & fileName
        Case 68
            Err.Description = "Dysk: " & driveName & " nie istnieje"
        Case 76
            Err.Description = "Nie ma takiego katalogu " & dirPath
        End Select
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Sprawdza czy podana œcie¿ka dostêpu do katalogu jest poprawna
' Wyœwietla informacje o b³êdach w œcie¿ce dostêpu
' @param String: thePath
' @return Boolean: Srue | False
Function DirPathValidate(thePath As String) As Boolean
    Const procName As String = "DirPathValidate(thePath)"
    Dim regex As New RegExp
    Dim pathState As String
    Dim Dirs() As String, count As Integer
    Dim theDir As String
    Dim isValid As Boolean
    Dim msg As String
    Dim i As Integer
    
On Error GoTo ErrorHandler
    regex.IgnoreCase = True
    ' Dysk sieciowy lub nie zmapowany udzia³
    If Left(thePath, 2) = "\\" Then pathState = "net"
    ' Dysk lokalny lub zmapowany udzia³ sieciowy
    regex.pattern = "^[a-z]:"
    If regex.Test(thePath) Then pathState = "local"
    ' Tworzy tablicê katalogów z podanej œcie¿ki
    Dirs = Split(thePath, "\")
    count = UBound(Dirs)
    Select Case pathState
    ' Dysk sieciowy lub nie zmapowany udzia³
    Case "net"
        theDir = "\\" & Dirs(2) & "\" & Dirs(3)
        ' Sprawdza poprawnoœæ udzia³u sieciowego
        isValid = DriveExists(theDir)
        If Not isValid Then
            msg = "Brak dostêpu lub b³êdna nazwa udzia³u sieciowego: " & _
                   vbNewLine & """" & theDir & """"
            GoTo NoValid
        End If
        ' Sprawdza poprawnoœæ kolejnych elementów œcie¿ki dostêpu
        For i = 4 To count
            theDir = theDir & "\" & Dirs(i)
            isValid = DirExists(theDir)
            If Not isValid Then
                msg = "Brak dostêpu lub b³êdna œcie¿ka dostêpu: " & _
                       vbNewLine & """" & theDir & """"
                GoTo NoValid
            End If
        Next i
    ' Dysk lokalny lub zmapowany udzia³ sieciowy
    Case "local"
        ' Sprawdza istnienie dysku
        theDir = Dirs(0)
        isValid = DriveExists(theDir)
        If Not isValid Then
            msg = "B³êdna nazwa dysku lub udzia³u sieciowego """ & theDir & """"
            GoTo NoValid
        End If
        ' Sprawdza poprawnoœæ kolejnych elementów œcie¿ki dostêpu
        For i = 1 To count
            theDir = theDir & "\" & Dirs(i)
            isValid = DirExists(theDir)
            If Not isValid Then
                msg = "B³êdna œcie¿ka dostêpu: " & _
                vbNewLine & """" & theDir & """"
                GoTo NoValid
            End If
        Next i
    Case Else
        msg = "B³êdna œcie¿ka dostêpu """ & thePath & """"
        GoTo NoValid
    End Select

DirPathValidate = True
Exit Function

NoValid:
MsgBox msg
DirPathValidate = False
Exit Function

ErrorHandler:
    If Not errState Then
        errState = True
        Err.Source = modName & procName & " linia " & Erl
    Else
        Err.Source = Err.Source & vbNewLine & _
                     modName & procName & " linia " & Erl
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Tworzy nowy katalog
' @param String dirName - nazwa nowego katalogu
' @param String dirPath - œcie¿ka dostêpu do nowego katalogu
' return Boolean: True | False
Public Function CreateDir(fullDirName As String) As Boolean
    Const procName As String = "CreateDir(fullDirName)"
    Dim fs, folder
    Dim parentDir As String
    
On Error GoTo ErrorHandler
    Set fs = CreateObject("Scripting.FileSystemObject")
    parentDir = fs.GetParentFolderName(fullDirName)
    
    If DirExists(parentDir) Then
       If Not DirExists(fullDirName) Then
            ' Zak³adam katalog
1           Set folder = fs.CreateFolder(fullDirName)
            CreateDir = True
        Else
            ' Katalog ju¿ istnieje
            CreateDir = True
        End If
    Else
        'Brak katalogu nadrzêdnego
        CreateDir = False
        GoTo Done
    End If
Done:
    Set fs = Nothing
    Set folder = Nothing
    Exit Function

ErrorHandler:
    If Not errState Then
        errState = True
        Err.Source = modName & procName & " linia " & Erl
    Else
        Err.Source = Err.Source & vbNewLine & _
                     modName & procName & " linia " & Erl
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function




' Usuwa katalog
' @param String fullDirName - pe³na œcie¿ka dostêpu do usuwanego katalogu
' return Boolean: True | vbNullString
Public Function DeleteDir(fullDirName As String) As Boolean
    Const procName As String = "DeleteDir(fullDirName)"
    Dim fs, folder
On Error GoTo ErrorHandler
    Set fs = CreateObject("Scripting.FileSystemObject")
1   Set folder = fs.GetFolder(fullDirName)
    fs.DeleteFolder (fullDirName)
    Set fs = Nothing
Done:
    DeleteDir = True
Exit Function

ErrorHandler:
    If Not errState Then
        errState = True
        Err.Source = modName & procName & " linia " & Erl
    Else
        Err.Source = Err.Source & vbNewLine & _
                     modName & procName & " linia " & Erl
    End If
    Select Case Err.Number
        Case 76
            Err.Description = "Nie ma takiego katalogu " & fullDirName
        End Select
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Usuwa plik
' @param String fullFileName - pe³na œcie¿ka dostêpu do pliku
' return Boolean: True | vbNullString
Public Function DeleteFile(fullFileName As String)
    Const procName As String = "DeleteFile(fullFileName)"
    Dim fs, file
    Dim fileName As String
On Error GoTo ErrorHandler
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    fileName = fs.GetFileName(fullFileName)
1   fs.DeleteFile (fullFileName)
    Set fs = Nothing
Done:
DeleteFile = True
Exit Function

ErrorHandler:
    If Not errState Then
        errState = True
        Err.Source = modName & procName & " linia " & Erl
    Else
        Err.Source = Err.Source & vbNewLine & _
                     modName & procName & " linia " & Erl
    End If
    Select Case Err.Number
        Case 53
            Err.Description = "Nie znaleziono pliku: " & fileName
        End Select
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Zwraca plik jako obiekt
' @param String fullFileName - pe³na œcie¿ka dostêpu do pliku
' return Object: File
Public Function GetFile(fullFileName As String) As Object
    Const procName As String = "GetFile(fullFileName)"
    Dim fs, file
    Dim fileName As String
On Error GoTo ErrorHandler
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    fileName = fs.GetFileName(fullFileName)
1   Set file = fs.GetFile(fullFileName)
    Set GetFile = file
    Set fs = Nothing
    Set file = Nothing
Done:
Exit Function

ErrorHandler:
    If Not errState Then
        errState = True
        Err.Source = modName & procName & " linia " & Erl
    Else
        Err.Source = Err.Source & vbNewLine & _
                     modName & procName & " linia " & Erl
    End If
    Select Case Err.Number
        Case 53
            Err.Description = "Nie znaleziono pliku: " & fileName
        End Select
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Zwraca zawartoœæ pliku tekstowego fileName jako objekt
' @param String fullFileName - pe³na œcie¿ka dostêpu do pliku
' return Object: TextStream
Public Function GetTextStream(fullFileName As String) As Object
    Const procName As String = "GetTextStream(fullFileName)"
    Dim fs, file
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    
    ' Format otwieranego pliku
    ' 0=TristateFalse - Otwiera plik jako ASCII. (default)
    ' 1=TristateTrue - Otwiera plik jako Unicode.
    ' 2=TristateUseDefault - Otwiera plik przy u¿yciu domyœlnej konfiguracji systemu.
    Const TristateFalse = 0, TristateTrue = 1, TristateUseDefault = 2

    Set fs = CreateObject("Scripting.FileSystemObject")
1   Set file = GetFile(fullFileName)
    Set GetTextStream = file.OpenAsTextStream(ForReading, TristateFalse)
    Set fs = Nothing
    Set file = Nothing
End Function




