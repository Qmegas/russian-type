VERSION 5.00
Begin VB.Form frmCustom 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Свои подстановки"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSear 
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtRepl 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Добавить"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Удалить"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.ListBox lList 
      Height          =   2010
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.CheckBox chkUse 
      Caption         =   "Использовать свои подстановки"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Находить:"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Заменять на:"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    If txtSear = vbNullString Then
        MsgBox "Поле поиска пусто", vbInformation, App.Title
        Exit Sub
    End If
    lList.AddItem txtSear & Chr(10) & txtRepl
    txtSear = vbNullString
    txtRepl = vbNullString
End Sub

Private Sub cmdDel_Click()
    If lList.ListIndex = -1 Then Exit Sub
    lList.RemoveItem lList.ListIndex
End Sub

Private Sub Form_Load()
    Dim j As Long
    
    If UseCustomRep Then chkUse.Value = 1
    For j = LBound(customRep) To UBound(customRep)
        If customRep(j).nSearch <> vbNullString Then
            lList.AddItem customRep(j).nSearch & Chr(10) & customRep(j).nReplace
        End If
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim j As Long, i As Long
    
    UseCustomRep = (chkUse.Value = 1)
    If lList.ListCount = 0 Then
        ReDim customRep(1 To 1)
    Else
        ReDim customRep(1 To lList.ListCount)
        For j = 1 To lList.ListCount
            i = InStr(1, lList.List(j - 1), Chr(10))
            customRep(j).nSearch = Mid(lList.List(j - 1), 1, i - 1)
            customRep(j).nReplace = Mid(lList.List(j - 1), i + 1)
        Next
    End If
End Sub
