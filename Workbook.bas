Attribute VB_Name = "Workbook"
Option Explicit

' Sprawdza czy skoroszyt istnieje
' @param string fullFileName - pełna ścieżka dostępu do skoroszytu.
' @return true | false
Function WoorkbookExists(ByVal fullFileName As String) As Boolean
    If Dir(fullFileName) = "" Then
        WoorkbookExists = False
    Else
        WoorkbookExists = True
    End If
End Function

' Sprawdza czy skoroszyt jest otwarty
' @param string fileName
' @return true | false
Function isWorkbookOpen(fileName) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Name = fileName Then
            isWorkbookOpen = True
            Exit Function
        Else
            isWorkbookOpen = False
        End If
    Next
End Function

' Zwraca otwarty skoroszyt
' Jeżeli jest zamknięty to go otwiera
' Jeżeli skoroszyt nie istnieje zwraca "Nothing"
' @param string fullFileName
' @return Workbook | Nothig
Function GetWorkbook(ByVal fullFileName As String) As Workbook
    Dim fileName As String
    fileName = Dir(fullFileName)
    On Error Resume Next
    Dim wb As Workbook
    Set wb = Workbooks(fileName)
    If wb Is Nothing Then
        Set wb = Workbooks.Open(fullFileName)
    End If
    On Error GoTo 0
    Set GetWorkbook = wb
End Function