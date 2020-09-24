VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Felicias' Diary"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7500
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7515
      Top             =   1080
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2055
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2040
      Width           =   6735
   End
   Begin VB.PictureBox pTitle1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      Picture         =   "frmMain.frx":C80E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   0
      Top             =   0
      Width           =   7500
      Begin VB.Label lTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Felicias' Diary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Label lDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   360
      Left            =   6000
      TabIndex        =   4
      Top             =   600
      Width           =   1410
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   2535
      Left            =   150
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   7215
   End
   Begin VB.Label lFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 16
Option Explicit
'Make Window Minimize
'Me.WindowState = 1
' mainForms load procedures.
Private Sub Form_Load()
' Check user settings to see if they want app ontop.
Dim msg As String
msg = Read_Ini("Windows", "pos")
If msg = 1 Then
AlwaysOnTop Me, True
frmMenu1.mnuOnTop.Checked = True
frmMenu1.mnuOffTop.Checked = False
End If
If msg = 0 Then
AlwaysOnTop Me, False
frmMenu1.mnuOffTop.Checked = True
frmMenu1.mnuOnTop.Checked = False
End If
' Make the shape of the form.
DoTransparency Me
' Set mnuSave to false, there is not file to save yet.
frmMenu1.mnuSave.Enabled = False
End Sub
' Change menus' color.
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lFile.ForeColor = &HFF80FF
End Sub
' Menu popup.
Private Sub lFile_Click()
frmMain.PopupMenu frmMenu1.mnuFile, 1
End Sub
' Change menus' color.
Private Sub lFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lFile.ForeColor = &HFFC0FF
End Sub
' Drag app. form the title bar.
Private Sub pTitle1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub
' Show the date.
Private Sub Timer1_Timer()
lDate.Caption = Format$(Now, "mm/dd/yy")
End Sub
