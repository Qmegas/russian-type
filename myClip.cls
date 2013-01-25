VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Private Const CF_UNICODETEXT As Long = 13
Private Const CF_TEXT = 1

Public Function getText() As String
    Dim myStrPtr As Long, myLen As Long, sBuffer As String, myLock As Long
    Dim useFormat As Long
    
    #If dbg = 1 Then
        On Error GoTo ErrorHandler
    #End If
    
    sBuffer = vbNullString
    
    OpenClipboard 0&
    If IsClipboardFormatAvailable(CF_UNICODETEXT) <> 0 Then
        myStrPtr = GetClipboardData(CF_UNICODETEXT)
        If myStrPtr <> 0 Then
            myLock = GlobalLock(myStrPtr)
            myLen = GlobalSize(myStrPtr)
            sBuffer = String$(myLen \ 2 - 1, vbNullChar)
            lstrcpy StrPtr(sBuffer), myLock
            GlobalUnlock myStrPtr
        End If
    Else
        If IsClipboardFormatAvailable(CF_TEXT) <> 0 Then
            myStrPtr = GetClipboardData(CF_TEXT)
            If myStrPtr <> 0 Then
                myLen = lstrlen(myStrPtr)
                If myLen > 0 Then
                    sBuffer = Space$(myLen)
                    CopyMemory ByVal sBuffer, ByVal myStrPtr, myLen
                End If
            End If
        End If
    End If
    CloseClipboard
    
    getText = sBuffer
    
    #If dbg = 1 Then
ErrorHandler:
        If Err.Number <> 0 Then
            MsgBox "Function: getText" & vbCrLf & "Error:" & Err.Number & vbCrLf & Err.Description
            Err.Clear
        End If
    #End If
End Function

Public Sub Clear()
    OpenClipboard 0&
    EmptyClipboard
    CloseClipboard
End Sub

Public Function SetText(OuStr) As Boolean
      Dim hData As Long
      Dim lpData As Long
      Dim Buffer() As Byte
      
      #If dbg = 1 Then
        On Error GoTo ErrorHandler
      #End If

      WorkingNow = True
      If OpenClipboard(0&) Then
         ' Convert data to ANSI byte array.
         Buffer = StrConv(OuStr & vbNullChar, vbFromUnicode)
         hData = GlobalAlloc(GMEM_FIXED, UBound(Buffer) + 1)
         If hData Then
            ' Copy data to alloc'd memory.
            lpData = GlobalLock(hData)
            Call CopyMemory(ByVal lpData, Buffer(0), UBound(Buffer) + 1)
            Call GlobalUnlock(hData)
            SetText = CBool(SetClipboardData(CF_TEXT, hData))
         End If

         ' Place Unicode text on clipboard, too.
         ' Not strictly necessary, as Windows will
         ' convert by default, with above code.
         ' Already null-terminated, so just
         ' allocate sufficient space for copy.
         hData = GlobalAlloc(GMEM_FIXED, LenB(OuStr) + 2)
         If hData Then
            ' Copy data to alloc'd memory.
            lpData = GlobalLock(hData)
            Call CopyMemory(ByVal lpData, ByVal StrPtr(OuStr), LenB(OuStr) + 2)
            Call GlobalUnlock(hData)
            ' Hand data off to clipboard
            Call SetClipboardData(CF_UNICODETEXT, hData)
         End If

         Call CloseClipboard
      End If
      WorkingNow = False
      
      #If dbg = 1 Then
ErrorHandler:
        If Err.Number <> 0 Then
            MsgBox "Function: SetText" & vbCrLf & "Error:" & Err.Number & vbCrLf & Err.Description
            Err.Clear
        End If
    #End If
End Function
