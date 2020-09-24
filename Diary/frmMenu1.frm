VERSION 5.00
Begin VB.Form frmMenu1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Save as"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuText 
         Caption         =   "&Text"
         Begin VB.Menu mnuCut 
            Caption         =   "&Cut"
         End
         Begin VB.Menu mnuSelect 
            Caption         =   "&Select All"
         End
         Begin VB.Menu mnuDelete 
            Caption         =   "&Delete"
         End
         Begin VB.Menu mnuCopy 
            Caption         =   "&Copy"
         End
         Begin VB.Menu mnuPaste 
            Caption         =   "&Paste"
         End
         Begin VB.Menu mnuReverse 
            Caption         =   "&Reverse text"
         End
      End
      Begin VB.Menu mnuSplit5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetup 
         Caption         =   "&Setup"
         Begin VB.Menu mnuOnTop 
            Caption         =   "&Make window ontop"
         End
         Begin VB.Menu mnuOffTop 
            Caption         =   "&Make window offtop"
         End
      End
      Begin VB.Menu mnuSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFelicia 
         Caption         =   "&Felicia"
         Begin VB.Menu mnuChangePW 
            Caption         =   "&Change your password"
         End
         Begin VB.Menu mnuLog 
            Caption         =   "&View your log"
         End
      End
      Begin VB.Menu mnuSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMin 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu mnuSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMenu1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Temp_File = "Temp.dry"
' Msg for changing pasword.
Const Msg_pChange = "Please type your new" & _
                  " user password. Password must be" & _
                  " NINE letters or numbers Long!"
' Applications title.
Const Apps_Title = "Felicias' Diary"
' Change user password.
Private Sub mnuChangePW_Click()
' This is the new password.
Dim New_Password As String
' Store the new passy in New_Password.
New_Password = InputBox(Msg_pChange, Apps_Title)
' If password=nothing goto Error Handler.
If New_Password = "" Then GoTo BumFuck
' Else if user entered a passy save it.
If DoPass(25, New_Password) = True Then
' Tell user password was stored without error.
MsgBox "Password is stored!", vbExclamation, Apps_Title
End If
Exit Sub
' Error. User did not enter a password.
BumFuck:
' Show a msg to inform user invalid password.
MsgBox "You did not enter a" & _
       " valid password." & vbCrLf & _
       "Please try again.", vbExclamation, Apps_Title
' Now exit the sub.
Exit Sub
End Sub
' Copy text to clipboard.
Private Sub mnuCopy_Click()
Clipboard.Clear
Clipboard.SetText frmMain.Text1.Text
End Sub
' Cut text and add to clipboard.
Private Sub mnuCut_Click()
Clipboard.Clear
Clipboard.SetText frmMain.Text1.SelText
frmMain.Text1.SelText = vbNullString
End Sub
' Delete text.
Private Sub mnuDelete_Click()
frmMain.Text1.SelText = vbNullString
End Sub
' Exit the application.
Private Sub mnuExit_Click()
End
End Sub
' Show the log.
Private Sub mnuLog_Click()
frmLog.Show
End Sub
' Minimize window.
Private Sub mnuMin_Click()
' Makes the Window Minimized in the taskbar.
frmMain.WindowState = 1
End Sub
' Take window off top.
Private Sub mnuOffTop_Click()
AlwaysOnTop frmMain, False
mnuOffTop.Checked = True
mnuOnTop.Checked = False
Write_Ini "Windows", "pos", "0"
End Sub
' Make window stay on top.
Private Sub mnuOnTop_Click()
AlwaysOnTop frmMain, True
mnuOnTop.Checked = True
mnuOffTop.Checked = False
Write_Ini "Windows", "pos", "1"
End Sub
' Open new file.
Private Sub mnuOpen_Click()
FileOpen
End Sub
' Paste text.
Private Sub mnuPaste_Click()
frmMain.Text1.SelText = Clipboard.GetText()
End Sub
' Print the Diary Contents.
Private Sub mnuPrint_Click()
' Set the printer up
Printer.FontSize = 18
' Print date time and text.
Printer.Print CStr(Date & " " & Time) & vbCrLf & frmMain.Text1.Text
' Close the document up.
Printer.EndDoc
End Sub
' Reverse text.
Private Sub mnuReverse_Click()
Dim msg As String
msg = StrReverse(frmMain.Text1.Text)
frmMain.Text1.Text = msg
End Sub
' Save new diary file.
Private Sub mnuSaveAs_Click()
SaveDfile
End Sub
' Select text.
Private Sub mnuSelect_Click()
frmMain.Text1.SelStart = 0
frmMain.Text1.SelLength = Len(frmMain.Text1.Text)
End Sub
