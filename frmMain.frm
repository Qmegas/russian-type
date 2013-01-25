VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   1335
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   1965
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   1965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu cmm 
      Caption         =   "cmmenu"
      Visible         =   0   'False
      Begin VB.Menu cms 
         Caption         =   "Отключить"
         Index           =   1
      End
      Begin VB.Menu cms 
         Caption         =   "Форматировать текст"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu cms 
         Caption         =   "Режим"
         Index           =   20
         Begin VB.Menu cmode 
            Caption         =   "Транслит Eng->Rus"
            Checked         =   -1  'True
            Index           =   1
            Shortcut        =   +^{F1}
         End
         Begin VB.Menu cmode 
            Caption         =   "Транслит Rus->Eng"
            Index           =   2
            Shortcut        =   +^{F2}
         End
         Begin VB.Menu cmode 
            Caption         =   "Перенабор Eng->Rus"
            Index           =   3
            Shortcut        =   +^{F3}
         End
         Begin VB.Menu cmode 
            Caption         =   "Перенабор Rus->Eng"
            Index           =   4
            Shortcut        =   +^{F4}
         End
         Begin VB.Menu cmode 
            Caption         =   "Транслит Eng->Heb (Experemental)"
            Index           =   5
            Shortcut        =   +^{F5}
         End
      End
      Begin VB.Menu cms 
         Caption         =   "История"
         Index           =   30
         Begin VB.Menu cmhist 
            Caption         =   "1"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu cmhist 
            Caption         =   "-"
            Index           =   998
         End
         Begin VB.Menu cmhist 
            Caption         =   "Очистить"
            Index           =   999
         End
      End
      Begin VB.Menu cms 
         Caption         =   "Свои подстоновки..."
         Index           =   35
      End
      Begin VB.Menu cms 
         Caption         =   "О Программе..."
         Index           =   40
      End
      Begin VB.Menu cms 
         Caption         =   "-"
         Index           =   50
      End
      Begin VB.Menu cms 
         Caption         =   "Выход"
         Index           =   60
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmhist_Click(Index As Integer)
    Dim scState As Boolean, i As Integer
    Dim j As Integer
    
    #If dbg = 1 Then
        On Error GoTo ErrorHandler
    #End If
    
    If Index = 999 Then
        j = HistData.Count
        Set HistData = Nothing
        Set HistData = New Collection
        For i = j To 1 Step -1
            Unload cmhist(i)
        Next
    Else
        scState = ScanClip
        ScanClip = False
        WorkingNow = True
        iClip.Clear
        iClip.SetText HistData.Item(Index)
        ScanClip = scState
        WorkingNow = False
    End If
    
    #If dbg = 1 Then
ErrorHandler:
        If Err.Number <> 0 Then
            MsgBox "Function: cmhist_Click" & vbCrLf & "Error:" & Err.Number & vbCrLf & Err.Description
            Err.Clear
        End If
    #End If
End Sub

Public Sub cmode_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To 4
        cmode(i).Checked = False
    Next
    cmode(Index).Checked = True
    SMode = Index
End Sub

Public Sub cms_Click(Index As Integer)
    #If dbg = 1 Then
        On Error GoTo ErrorHandler
    #End If
    
    Select Case Index
        Case 1 'Disable
            If cms(1).Caption = "Отключить" Then
                ScanClip = False
                WorkingNow = True
                iClip.Clear
                iClip.SetText MyString
                WorkingNow = False
                cms(1).Caption = "Включить"
                SetTrayIcon NIM_MODIFY, hwnd, dPiksa.Handle, App.Title
            Else
                WorkingNow = True
                SetClipboardViewer Me.hwnd
                WorkingNow = False
                If (MyString <> vbNullString) And (Not IsExeption) Then
                    RunString
                    iClip.Clear
                    iClip.SetText OuStr
                End If
                SetTrayIcon NIM_MODIFY, hwnd, ePiksa.Handle, App.Title
                ScanClip = True
                cms(1).Caption = "Отключить"
            End If
            cms(5).Enabled = Not ScanClip
        Case 5 'Format text
            cms(5).Checked = Not cms(5).Checked
        Case 35
            frmCustom.Show
        Case 40 'About
            frmAbout.Show
        Case 60 'Exit
            Unload Me
    End Select
    
    #If dbg = 1 Then
ErrorHandler:
        If Err.Number <> 0 Then
            MsgBox "Function: cms_Click" & vbCrLf & "Error:" & Err.Number & vbCrLf & Err.Description
            Err.Clear
        End If
    #End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    #If dbg = 1 Then
        On Error GoTo ErrorHandler
    #End If
    
    SetTrayIcon NIM_ADD, hwnd, ePiksa.Handle, App.Title
    
    RegisterHotKey hwnd, 1, MOD_CONTROL Or MOD_SHIFT, vbKeyF1
    RegisterHotKey hwnd, 2, MOD_CONTROL Or MOD_SHIFT, vbKeyF2
    RegisterHotKey hwnd, 3, MOD_CONTROL Or MOD_SHIFT, vbKeyF3
    RegisterHotKey hwnd, 4, MOD_CONTROL Or MOD_SHIFT, vbKeyF4
    RegisterHotKey hwnd, 5, MOD_CONTROL Or MOD_SHIFT, vbKeyDown
    RegisterHotKey hwnd, 6, MOD_CONTROL Or MOD_SHIFT, 192  'Ctrl+Shift+`
    
    prev = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WinProc)
    SetClipboardViewer Me.hwnd
    
    #If dbg = 1 Then
ErrorHandler:
        If Err.Number <> 0 Then
            MsgBox "Function: Form_Load" & vbCrLf & "Error:" & Err.Number & vbCrLf & Err.Description
            Err.Clear
        End If
    #End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    #If dbg = 1 Then
        On Error GoTo ErrorHandler
    #End If
    
    SetTrayIcon NIM_DELETE, hwnd, 0, vbNullString
    Set ePiksa = Nothing
    Set dPiksa = Nothing
    UnregisterHotKey hwnd, 1
    UnregisterHotKey hwnd, 2
    UnregisterHotKey hwnd, 3
    UnregisterHotKey hwnd, 4
    UnregisterHotKey hwnd, 5
    UnregisterHotKey hwnd, 6
    
    #If dbg = 1 Then
ErrorHandler:
        If Err.Number <> 0 Then
            MsgBox "Function: Form_QueryUnload" & vbCrLf & "Error:" & Err.Number & vbCrLf & Err.Description
            Err.Clear
        End If
    #End If
End Sub

Public Sub LoadHistory()
    Dim i As Integer, tmp As String
    On Error Resume Next
    
    For i = 1 To HistData.Count
        tmp = HistData.Item(i)
        If Len(tmp) > 25 Then cmhist(i).Caption = Left(tmp, 25) & "..." _
            Else cmhist(i).Caption = tmp
    Next
    
    #If dbg = 1 Then
        If Err.Number <> 0 Then
            MsgBox "Function: Form_QueryUnload" & vbCrLf & "Error:" & Err.Number & vbCrLf & Err.Description
        End If
    #End If
    
    Err.Clear
End Sub
