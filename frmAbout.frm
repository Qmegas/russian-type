VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "О программе..."
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check5 
      Caption         =   "Русско-англиский словарь 2"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Вести историю использования буффера"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   3495
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Использовать символ ` для приостановки работы"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   4215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Поменять Ш и Щ"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Использовать '4' как 'Ч'"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Columns         =   5
      Height          =   1815
      ItemData        =   "frmAbout.frx":0000
      Left            =   120
      List            =   "frmAbout.frx":0067
      TabIndex        =   3
      Top             =   1440
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Закрыть"
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "http://qmegas.info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Created by Megas © (2004-2007)"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Версия 2.60"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Russian Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If ch4b Then Check1.Value = 1
    If SwapSW Then Check2.Value = 1
    If UsePauseMode Then Check3.Value = 1
    If UseHistory Then Check4.Value = 1
    If RESlovar2 Then Check5.Value = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ch4b = (Check1.Value = 1)
    SwapSW = (Check2.Value = 1)
    UsePauseMode = (Check3.Value = 1)
    UseHistory = (Check4.Value = 1)
    RESlovar2 = (Check5.Value = 1)
End Sub

Private Sub Label5_Click()
    ShellExecute Me.hwnd, "open", "http://qmegas.info", vbNullString, App.Path, vbNormalFocus
End Sub
