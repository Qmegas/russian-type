Attribute VB_Name = "Module1"
Option Explicit

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Type CUSTREP
    nSearch As String
    nReplace As String
End Type

Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long
Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Const WM_USER = &H400
Public Const TRAY_CALLBACK = (WM_USER + 101&)
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIM_MODIFY = &H1
Public Const NIM_ADD = &H0&
Public Const NIM_DELETE = &H2&
Public Const GWL_WNDPROC = (-4&)
Public Const WM_DRAWCLIPBOARD = &H308
Public Const CF_TEXT = 1
Public Const CF_UNICODETEXT = 13
Public Const GMEM_FIXED = 0
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const MAX_HIST_SIZE = 30
Private Const WM_HOTKEY = &H312
Private Const WM_ACTIVATEAPP = &H1C


Public prev As Long
Public MyString As String
Public ScanClip As Boolean
Public WorkingNow As Boolean
Public CurPz As Long
Public OuStr As String
Public b As String
Public SMode As Integer
Public UsePauseMode As Boolean
Public OnPauseNow As Boolean
Public RESlovar2 As Boolean

Public customRep() As CUSTREP
Public UseCustomRep As Boolean

Public ch4b As Boolean
Public SwapSW As Boolean
Public UseHistory As Boolean
Public ePiksa As Picture
Public dPiksa As Picture
Public HistData As New Collection

Public iClip As New myClip

Sub Main()
    If App.PrevInstance Then End
    
    ch4b = True
    ScanClip = True
    SwapSW = False
    UsePauseMode = True
    UseHistory = True
    WorkingNow = False
    RESlovar2 = False
    UseCustomRep = False
    ReDim customRep(1 To 1)
    SMode = 1
    Set ePiksa = LoadResPicture(1, vbResIcon)
    Set dPiksa = LoadResPicture(2, vbResIcon)
    Load frmMain
End Sub

Public Function WinProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim errStick As Integer
    #If dbg = 1 Then
        On Error GoTo ErrorHandler
    #End If
    
    WinProc = CallWindowProc(prev, hwnd, Msg, wParam, lParam)
    Select Case Msg
        Case TRAY_CALLBACK 'Tray icon
            If lParam = 517 Then
                With frmMain
                    .cms(30).Visible = UseHistory
                    If UseHistory Then .LoadHistory
                    SetForegroundWindow frmMain.hwnd
                    .PopupMenu .cmm, , , , .cms(40)
                End With
            End If
            '== Debug ==
            errStick = 1
            '===========
            If lParam = 513 Then 'LBDown
                ScanClip = False
                WorkingNow = True
                iClip.Clear
                iClip.SetText MyString
                WorkingNow = False
                DoEvents
                If frmMain.cms(1).Caption = "Отключить" _
                    Then ScanClip = True
            End If
        Case WM_DRAWCLIPBOARD 'Clipboard
            If WorkingNow Then Exit Function
            
            '== Debug ==
            errStick = 2
            '===========
            If UseHistory Then
                MyString = iClip.getText
                If Len(MyString) <> 0 Then
                    If HistData.Count > 0 Then HistData.Add MyString, , 1 _
                        Else HistData.Add MyString
                    If HistData.Count > MAX_HIST_SIZE Then
                        HistData.Remove MAX_HIST_SIZE + 1
                    Else
                        Load frmMain.cmhist(HistData.Count)
                        frmMain.cmhist(HistData.Count).Visible = True
                    End If
                End If
            End If
            
            '== Debug ==
            errStick = 3
            '===========
            If ScanClip Then
                ScanClip = False
                WorkingNow = True
                MyString = iClip.getText
                If (Len(MyString) <> 0) And (Not IsExeption) Then
                    RunString
                    If UseCustomRep Then RunCustom
                    '== Debug ==
                    errStick = 5
                    '===========
                    iClip.Clear
                    iClip.SetText OuStr
                End If
                ScanClip = True
                WorkingNow = False
            Else
                '== Debug ==
                errStick = 4
                '===========
                If frmMain.cms(5).Checked Then
                    WorkingNow = True
                    If (Len(iClip.getText) <> 0) Then
                        MyString = iClip.getText
                        OuStr = MyString
                        If UseCustomRep Then RunCustom
                        iClip.Clear
                        iClip.SetText OuStr
                    End If
                    WorkingNow = False
                End If
            End If
        Case WM_HOTKEY
            Select Case wParam
                Case Is < 5: frmMain.cmode_Click CInt(wParam)
                Case 5: frmMain.cms_Click 1
                Case 6: RefreshSysIcon
            End Select
    End Select
    
    #If dbg = 1 Then
ErrorHandler:
        If Err.Number <> 0 Then
            MsgBox "Function: WinProc" & vbCrLf & "Error:" & Err.Number & vbCrLf & Err.Description & vbCrLf & "Tag: " & CStr(errStick)
            Err.Clear
        End If
    #End If
End Function

Public Sub RunString()
    Dim tb As String
    
    #If dbg = 1 Then
        On Error GoTo ErrorHandler
    #End If
    
    OuStr = vbNullString
    CurPz = 1
    OnPauseNow = False

    Do While CurPz <= Len(MyString)
        tb = Mid(MyString, CurPz, 1)
        If (tb = "`") And (UsePauseMode) Then
            OnPauseNow = Not OnPauseNow
            tb = ""
        End If
        If Not OnPauseNow Then
            Select Case SMode
                Case 1: If RESlovar2 Then GetBukER_2 LCase(tb) Else GetBukER LCase(tb)
                Case 2: GetBukRE LCase(tb)
                Case 3: PerenaborER LCase(tb)
                Case 4: PerenaborRE LCase(tb)
                Case 5: GetBukEH (tb) 'Experemental mode
            End Select
        End If
        If LCase(tb) <> tb Then
            If Len(b) > 1 Then
                b = UCase(Left(b, 1)) & Right(b, Len(b) - 1)
            Else
                b = UCase(b)
            End If
        End If
        If OnPauseNow Then OuStr = OuStr & tb _
            Else OuStr = OuStr & b
        SkipOne
    Loop
    
    #If dbg = 1 Then
ErrorHandler:
        If Err.Number <> 0 Then
            MsgBox "Function: RunString" & vbCrLf & "Error:" & Err.Number & vbCrLf & Err.Description
            Err.Clear
        End If
    #End If
End Sub

Private Sub GetBukER(ib As String)
    Dim tb As String * 1, i As Long, j As Long
    
    Select Case ib
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
            i = InStr(CurPz, MyString, " ")
            If i = 0 Then
                b = Mid(MyString, CurPz, Len(MyString) - CurPz + 1)
            Else
                b = Mid(MyString, CurPz, i - CurPz)
            End If
            If (IsNumeric(b)) And (b <> "4") And (b <> "4,") Then
                CurPz = CurPz + Len(b) - 1
            Else
                If ch4b And (ib = "4") Then b = "ч" Else b = ib
            End If
        Case "a": b = "а"
        Case "b": b = "б"
        Case "c"
            b = "ц"
            If Chr(ViewNext) = "h" Then
                b = "ч"
                SkipOne
            End If
        Case "d":            b = "д"
        Case "e":            b = "е"
        Case "f"
            If LCase(Mid(MyString, CurPz, 4)) = "ftp:" Then
                i = InStr(CurPz, MyString, " ")
                If i = 0 Then
                    b = Mid(MyString, CurPz, Len(MyString) - CurPz + 1)
                    CurPz = Len(MyString)
                Else
                    b = Mid(MyString, CurPz, i - CurPz)
                    CurPz = i - 1
                End If
            Else
                b = "ф"
            End If
        Case "g":            b = "г"
        Case "h"
            If LCase(Mid(MyString, CurPz, 4)) = "http" Then
                i = InStr(CurPz, MyString, " ")
                j = InStr(CurPz, MyString, vbCrLf)
                If (j < i) And (j <> 0) Then i = j
                
                If i = 0 Then
                    b = Mid(MyString, CurPz, Len(MyString) - CurPz + 1)
                    CurPz = Len(MyString)
                Else
                    b = Mid(MyString, CurPz, i - CurPz)
                    CurPz = i - 1
                End If
            Else
                b = "х"
            End If
        Case "i":            b = "и"
        Case "j":            b = "ж"
        Case "k":            b = "к"
        Case "l":            b = "л"
        Case "m":            b = "м"
        Case "n":            b = "н"
        Case "o":            b = "о"
        Case "p":            b = "п"
        Case "r":            b = "р"
        Case "s"
            b = "с"
            If Chr(ViewNext) = "h" Then
                If SwapSW Then b = "щ" Else b = "ш"
                SkipOne
            End If
        Case "t":            b = "т"
        Case "u":            b = "у"
        Case "v":            b = "в"
        Case "w"
            If LCase(Mid(MyString, CurPz, 4)) = "www." Then
                i = InStr(CurPz, MyString, " ")
                If i = 0 Then
                    b = Mid(MyString, CurPz, Len(MyString) - CurPz + 1)
                    CurPz = Len(MyString)
                Else
                    b = Mid(MyString, CurPz, i - CurPz)
                    CurPz = i - 1
                End If
            Else
                If SwapSW Then b = "ш" Else b = "щ"
            End If
        Case "x":            b = "х"
        Case "y"
            
            tb = Chr(ViewNext)
            Select Case tb
                Case "a"
                    b = "я"
                    SkipOne
                Case "e"
                    b = "э"
                    SkipOne
                Case "i"
                    b = "ы"
                    SkipOne
                Case "o"
                    b = "ё"
                    SkipOne
                Case "u"
                    b = "ю"
                    SkipOne
                Case Else
                    b = "й"
            End Select
        Case "z":            b = "з"
        Case "'"
            b = "ь"
            If Chr(ViewNext) = "'" Then
                b = "ъ"
                SkipOne
            End If
        Case "^": b = "Ь"
        Case Else:            b = ib
    End Select
End Sub

Private Sub GetBukER_2(ib As String)
    Dim tb As String * 1, i As Integer
    
    Select Case ib
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
            i = InStr(CurPz, MyString, " ")
            If i = 0 Then
                b = Mid(MyString, CurPz, Len(MyString) - CurPz + 1)
            Else
                b = Mid(MyString, CurPz, i - CurPz)
            End If
            If (IsNumeric(b)) And (b <> "4") And (b <> "4,") Then
                CurPz = CurPz + Len(b) - 1
            Else
                If ch4b And (ib = "4") Then b = "ч" Else b = ib
            End If
        Case "a": b = "а"
        Case "b": b = "б"
        Case "c"
            b = "ц"
            If Chr(ViewNext) = "h" Then
                b = "ч"
                SkipOne
            End If
        Case "d":            b = "д"
        Case "e"
            b = "е"
            If Chr(ViewNext) = "h" Then
                b = "э"
                SkipOne
            End If
        Case "f"
            If LCase(Mid(MyString, CurPz, 4)) = "ftp:" Then
                i = InStr(CurPz, MyString, " ")
                If i = 0 Then
                    b = Mid(MyString, CurPz, Len(MyString) - CurPz + 1)
                    CurPz = Len(MyString)
                Else
                    b = Mid(MyString, CurPz, i - CurPz)
                    CurPz = i - 1
                End If
            Else
                b = "ф"
            End If
        Case "g"
            b = "г"
            If Chr(ViewNext) = "h" Then
                b = "ж"
                SkipOne
            End If
        Case "h"
            If LCase(Mid(MyString, CurPz, 4)) = "http" Then
                i = InStr(CurPz, MyString, " ")
                If i = 0 Then
                    b = Mid(MyString, CurPz, Len(MyString) - CurPz + 1)
                    CurPz = Len(MyString)
                Else
                    b = Mid(MyString, CurPz, i - CurPz)
                    CurPz = i - 1
                End If
            Else
                b = "х"
            End If
        Case "i":            b = "и"
        Case "j":            b = "й"
        Case "k":            b = "к"
        Case "l":            b = "л"
        Case "m":            b = "м"
        Case "n":            b = "н"
        Case "o":            b = "о"
        Case "p":            b = "п"
        Case "r":            b = "р"
        Case "s"
            b = "с"
            If Chr(ViewNext) = "h" Then
                If SwapSW Then b = "щ" Else b = "ш"
                SkipOne
            End If
        Case "t":            b = "т"
        Case "u":            b = "у"
        Case "v":            b = "в"
        Case "w"
            If LCase(Mid(MyString, CurPz, 4)) = "www." Then
                i = InStr(CurPz, MyString, " ")
                If i = 0 Then
                    b = Mid(MyString, CurPz, Len(MyString) - CurPz + 1)
                    CurPz = Len(MyString)
                Else
                    b = Mid(MyString, CurPz, i - CurPz)
                    CurPz = i - 1
                End If
            Else
                If SwapSW Then b = "ш" Else b = "щ"
            End If
        Case "x":            b = "х"
        Case "y"
            
            tb = Chr(ViewNext)
            Select Case tb
                Case "o"
                    b = "ё"
                    SkipOne
                Case "u"
                    b = "ю"
                    SkipOne
                Case "a"
                    b = "я"
                    SkipOne
                Case Else
                    b = "ы"
            End Select
        Case "z":            b = "з"
        Case "'"
            b = "ь"
            If Chr(ViewNext) = "'" Then
                b = "ъ"
                SkipOne
            End If
        Case Else:            b = ib
    End Select
End Sub

Private Sub GetBukRE(ib As String)
    Select Case ib
        Case "а": b = "a"
        Case "б": b = "b"
        Case "в": b = "v"
        Case "г": b = "g"
        Case "д": b = "d"
        Case "е": b = "e"
        Case "ё": b = "yo"
        Case "ж": b = "j"
        Case "з": b = "z"
        Case "и": b = "i"
        Case "й": b = "y"
        Case "к": b = "k"
        Case "л": b = "l"
        Case "м": b = "m"
        Case "н": b = "n"
        Case "о": b = "o"
        Case "п": b = "p"
        Case "р": b = "r"
        Case "с": b = "s"
        Case "т": b = "t"
        Case "у": b = "u"
        Case "ф": b = "f"
        Case "х": b = "x"
        Case "ц": b = "c"
        Case "ч": b = "ch"
        Case "ш": b = "sh"
        Case "щ": b = "w"
        Case "ь": b = "'"
        Case "ы": b = "yi"
        Case "ъ": b = "''"
        Case "э": b = "ye"
        Case "ю": b = "yu"
        Case "я": b = "ya"
        Case Else: b = ib
    End Select
End Sub

Private Sub PerenaborER(ib As String)
    Select Case ib
        Case "q": b = "й"
        Case "w": b = "ц"
        Case "e": b = "у"
        Case "r": b = "к"
        Case "t": b = "е"
        Case "y": b = "н"
        Case "u": b = "г"
        Case "i": b = "ш"
        Case "o": b = "щ"
        Case "p": b = "з"
        Case "[": b = "х"
        Case "{": b = "Х"
        Case "]": b = "ъ"
        Case "}": b = "Ъ"
        Case "a": b = "ф"
        Case "s": b = "ы"
        Case "d": b = "в"
        Case "f": b = "а"
        Case "g": b = "п"
        Case "h": b = "р"
        Case "j": b = "о"
        Case "k": b = "л"
        Case "l": b = "д"
        Case ";": b = "ж"
        Case ":": b = "Ж"
        Case "'": b = "э"
        Case Chr(34): b = "Э"
        Case "z": b = "я"
        Case "x": b = "ч"
        Case "c": b = "с"
        Case "v": b = "м"
        Case "b": b = "и"
        Case "n": b = "т"
        Case "m": b = "ь"
        Case ",": b = "б"
        Case "<": b = "Б"
        Case ".": b = "ю"
        Case ">": b = "Ю"
        Case "/": b = "."
        Case "?": b = ","
        Case Else: b = ib
    End Select
End Sub

Private Sub PerenaborRE(ib As String)
    Select Case ib
        Case "ё": b = "`"
        Case "й": b = "q"
        Case "ц": b = "w"
        Case "у": b = "e"
        Case "к": b = "r"
        Case "е": b = "t"
        Case "н": b = "y"
        Case "г": b = "u"
        Case "ш": b = "i"
        Case "щ": b = "o"
        Case "з": b = "p"
        Case "х": b = "["
        Case "ъ": b = "]"
        Case "ф": b = "a"
        Case "ы": b = "s"
        Case "в": b = "d"
        Case "а": b = "f"
        Case "п": b = "g"
        Case "р": b = "h"
        Case "о": b = "j"
        Case "л": b = "k"
        Case "д": b = "l"
        Case "ж": b = ";"
        Case "э": b = "'"
        Case "я": b = "z"
        Case "ч": b = "x"
        Case "с": b = "c"
        Case "м": b = "v"
        Case "и": b = "b"
        Case "т": b = "n"
        Case "ь": b = "m"
        Case "б": b = ","
        Case "ю": b = "."
        Case ".": b = "/"
        Case ",": b = "?"
        Case Else: b = ib
    End Select
End Sub

Private Sub GetBukEH(ib As String)
    Select Case ib
        Case "a": b = ChrW$(1488) 'Alef
        Case "b": b = ChrW$(1489) 'Bet
        Case "g": b = ChrW$(1490) 'Gimel
        Case "d": b = ChrW$(1491) 'Dalet
        Case "x": b = ChrW$(1492) 'Hei
        Case "v", "u", "o": b = ChrW$(1493) 'Vav
        Case "z": b = ChrW$(1494) 'Zain
        Case "h": b = ChrW$(1495) 'Het
        Case "T": b = ChrW$(1496) 'Tet
        Case "i": b = ChrW$(1497) 'Iud
        Case "K", "H" 'Xaf
            b = ChrW$(1499)
            If Chr(ViewNext) = " " Then b = ChrW$(1498)
        Case "l": b = ChrW(1500) 'lamed
        Case "m" 'Mem
            b = ChrW$(1502)
            If Chr(ViewNext) = " " Then b = ChrW$(1501)
        Case "n" 'Nun
            b = ChrW$(1504)
            If Chr(ViewNext) = " " Then b = ChrW$(1503)
        Case "s" 'Sin, shin
            b = ChrW$(1505)
            If Chr(ViewNext) = "h" Then
                b = ChrW$(1513)
                SkipOne
            End If
        Case "A": b = ChrW$(1506) 'Ain
        Case "p", "f" 'Pei, fei
            b = ChrW$(1508)
            If Chr(ViewNext) = " " Then b = ChrW$(1507)
        Case "c" 'Tzadik
            b = ChrW$(1510)
            If Chr(ViewNext) = " " Then b = ChrW$(1509)
        Case "k": b = ChrW$(1511) 'Kuf
        Case "r": b = ChrW$(1512) 'Reish
        Case "t": b = ChrW$(1514) 'Tav
        Case "e": b = vbNullString
        Case Else: b = ib
    End Select
End Sub

Private Function ViewNext() As Integer
    Dim tb As String * 1
    tb = Mid(MyString, CurPz + 1, 1)
    If tb = "" Then tb = " "
    ViewNext = Asc(LCase(tb))
End Function

Private Function SkipOne()
    CurPz = CurPz + 1
End Function

Public Function SetTrayIcon(Mode As Long, hwnd As Long, Icon As Long, tip As String) As Long
    Dim nidTemp As NOTIFYICONDATA
    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hwnd = hwnd
    nidTemp.uID = 0&
    nidTemp.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nidTemp.uCallbackMessage = TRAY_CALLBACK
    nidTemp.hIcon = Icon
    nidTemp.szTip = tip & Chr$(0)
    SetTrayIcon = Shell_NotifyIconA(Mode, nidTemp)
End Function

Public Function IsExeption() As Boolean
    Dim tmp As String
    
    tmp = LCase(Trim(MyString))
    IsExeption = False
    IsExeption = IsNumeric(MyString)
    If Mid(tmp, 2, 2) = ":\" Then IsExeption = True
End Function

Sub RefreshSysIcon()
    SetTrayIcon NIM_DELETE, frmMain.hwnd, 0, vbNullString
    SetTrayIcon NIM_ADD, frmMain.hwnd, ePiksa.Handle, App.Title
    If Not ScanClip Then _
        SetTrayIcon NIM_MODIFY, frmMain.hwnd, dPiksa.Handle, App.Title
End Sub

Private Sub str_replace(ByVal Find As String, ByVal Replace As String, ByRef Expression As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare)
    On Error Resume Next
    Dim l As Long
    Dim lenR As Long
    Dim p1 As Long
    Dim p2 As Long
    Dim p21 As Long
    Dim s As String
    
    l = Len(Find)
    If (l = 0) Then Exit Sub
    
    lenR = Len(Replace)
    If (lenR = l) Then
        p1 = 1
        p2 = InStr(p1, Expression, Find, Compare)
        Do While (p2)
            Mid$(Expression, p1) = Mid$(Expression, p1, p2 - p1)
            Mid$(Expression, p2) = Replace
            p1 = p2 + l
            p2 = InStr(p1, Expression, Find, Compare)
        Loop
        Exit Sub
        
    ElseIf (lenR > l) Then
        s = Space$(Len(Expression) + (Len(Expression) \ l) * (lenR - l))
        
    Else
        s = Space$(Len(Expression))
    End If
    
    p21 = 1
    p1 = 1
    p2 = InStr(p1, Expression, Find, Compare)
    Do While (p2)
        Mid$(s, p21) = Mid(Expression, p1, p2 - p1)
        p21 = p21 + p2 - p1
        Mid$(s, p21) = Replace
        p21 = p21 + lenR
        p1 = p2 + l
        p2 = InStr(p1, Expression, Find, Compare)
    Loop
    Mid$(s, p21) = Mid$(Expression, p1)
    p21 = p21 + Len(Mid$(Expression, p1))
    s = Left$(s, p21 - 1)
    Expression = s
End Sub

Private Sub RunCustom()
    Dim j As Long
    
    For j = LBound(customRep) To UBound(customRep)
        str_replace customRep(j).nSearch, customRep(j).nReplace, OuStr
    Next
End Sub
