Attribute VB_Name = "modAddFiles"
Option Explicit

'This is one of the most important modules
'this will take all of the files that you
'select, to add into the playlist, and will
'add, one by one, each file you select in the
'common dialog.  This is very important for
'the multiple selection feature.

Public Declare Function ShowCursor& Lib "user32" (ByVal bShow As Boolean)
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Function CountFilesInList(ByVal FileList As String) As Integer
    Dim iCount As Integer
    Dim iPos As Integer

    iCount = 0
    For iPos = 1 To Len(FileList)
        If Mid$(FileList, iPos, 1) = Chr$(0) Then iCount = iCount + 1
    Next
    If iCount = 0 Then iCount = 1
    CountFilesInList = iCount
End Function

Function GetFileFromList(ByVal FileList As String, FileNumber As Integer) As String
    Dim iPos                As Integer
    Dim iCount              As Integer
    Dim iFileNumberStart    As Integer
    Dim iFileNumberLen      As Integer
    Dim sPath               As String

    If InStr(FileList, Chr(0)) = 0 Then
        GetFileFromList = FileList
    Else
        iCount = 0
        sPath = left(FileList, InStr(FileList, Chr(0)) - 1)
        If Right(sPath, 1) <> "\" Then sPath = sPath + "\"
        FileList = FileList + Chr(0)
        For iPos = 1 To Len(FileList)
            If Mid$(FileList, iPos, 1) = Chr(0) Then
                iCount = iCount + 1
                Select Case iCount
                    Case FileNumber
                        iFileNumberStart = iPos + 1
                    Case FileNumber + 1
                        iFileNumberLen = iPos - iFileNumberStart
                        Exit For
                End Select
            End If
        Next
        GetFileFromList = sPath + Mid(FileList, iFileNumberStart, iFileNumberLen)
    End If
End Function


