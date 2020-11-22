Attribute VB_Name = "Clipboard"
Option Explicit
#If VBA7 Then

Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function EmptyClipboard Lib "User32" () As LongPtr
Private Declare PtrSafe Function CloseClipboard Lib "User32" () As LongPtr
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "User32" (ByVal wFormat As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal wFormat As LongPtr) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As LongPtr, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32.dll" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr

#Else

Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long

#End If

Public Sub SetClipboard(sUniText As String)

#If VBA7 Then
    Dim iStrPtr As LongPtr
    Dim iLock As LongPtr
#Else
    Dim iStrPtr As Long
    Dim iLock As Long
#End If

    Dim iLen As Long

    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = &HD

    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE + GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    lstrcpy iLock, StrPtr(sUniText)
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Sub

Public Function GetClipboard() As String
#If VBA7 Then
    Dim iStrPtr As LongPtr
    Dim iLock As LongPtr
#Else
    Dim iStrPtr As Long
    Dim iLock As Long
#End If
    Dim iLen As Long
    Dim sUniText As String

    Const CF_UNICODETEXT As Long = 13&

    OpenClipboard 0&

    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        If iStrPtr Then
            iLock = GlobalLock(iStrPtr)
            iLen = GlobalSize(iStrPtr)
            sUniText = String$(iLen \ 2& - 1&, vbNullChar)
            lstrcpy StrPtr(sUniText), iLock
            GlobalUnlock iStrPtr
        End If
        GetClipboard = sUniText
    End If

    CloseClipboard
End Function

