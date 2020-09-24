VERSION 5.00
Begin VB.Form frmCheck 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' P                                                         P
'   E                                                   E
'       A                                           A
'           N                                   N
'               U                           U
'                   T      is good      T
'                   B        4U         B
'               U                           U
'           T                                   T
'       T                                           T
'   E                                                   E
' R                                                         R
' Flag for checking right passy.
Dim ShitAccepted As Boolean
' Data entered for var, by user
Dim UserData As String
' Felicias' stored password
Dim FeliciaData As String
' Msg for the InputBox
Const Msg_User = "" & _
    "You must enter your user" & _
    " password to gain acces" & _
    " to Felicias' diary." & vbCrLf & _
    "If you do not enter a valid" & _
    " password then this program" & _
    " will auto-shutdown!" & vbCrLf & _
    "Thank you and have a fine day!" & vbCrLf & _
    "Note:" & vbCrLf & _
    "     Password is case SENSITIVE." & vbCrLf & _
    "By entering a password you agree that" & vbCrLf & _
    " you are infact Felicia Abrams."
' First warning msg user will see
Const Warn_Msg = "" & _
    "By clicking on yes" & _
    " you agree" & _
    " that you are Felicia Abrams." & vbCrLf & _
    "If however you are not" & _
    " Felicia Abrams then please" & _
    " press the NO button" & _
    " NOW!" & vbCrLf & _
    "Logging has begun..."
' Name of this application for InputBox
Const Apps_Title = "Felicais' diary"
' Pocedures when the form loads.
Private Sub Form_Load()
On Error GoTo AuthorsBlondeMoment
ShitAccepted = False
' Don't need the form
Me.Hide
' Warning user to exit if not Felicia.
If MsgBox(Warn_Msg, vbYesNo, Apps_Title) = vbNo Then
' Asshole alert now logem
DoLogg
' Don't allow this person to stroke my EGO!
End
Else
' Kewl It's Felicia, or so it seems.
GoTo CheckDaPassyYo
End If
' Do it
CheckDaPassyYo:
' Store users' entered data in UserData.
UserData = InputBox(Msg_User, Apps_Title)
' So they think there in huh, lets cross referance
' that shit.
FeliciaData = GetPass(25)
' Ok! User made it this far has to be Felicia. :)
If FeliciaData = UserData Then
' Show main Form.
frmMain.Show
' Damn! shit's accepted. :(
ShitAccepted = True
' Unload this form and load the frmMain
Unload Me
Exit Sub
Else: If ShitAccepted = False Then GoTo Check4BlondeMoment
' If they did not enter a passy or _
  just are inflicted with leopracy _
  lets give this Yoyo another chance.
End If
' Check for a blond moment.
Check4BlondeMoment:
' This is the beginning of a beautifull blond moment.
' Tell user what he/she/HeShe/It has done.
    If MsgBox("You did not enter" & _
     " a valid password!", vbAbortRetryIgnore, "" & _
     Apps_Title) = vbRetry Then
' Ok! Dis lil bitch thinks they can crack muh shit.
' So lets see if they get it right this time.
GoTo CheckDaPassyYo
        Else
' If user clicks Abort or Ignore they will end Application.
' Also loggem.
        DoLogg
            End
     End If
' I don't stroke my own Ego all the time.
' Only when I'm Upset. :)
' Now this is my lil Error handler.
AuthorsBlondeMoment:
    MsgBox "Sorry the Author of this" & _
           " program had a blonde moment" & _
           " and would like to resolve this" & _
           " matter with you over some Chocolate" & _
           " Milk." & vbCrLf & _
           "By pressing the Ok button at the bottom" & _
           " you forgive the Author of whatever" & _
           " trauma that has happened." & vbCrLf & _
           "Please take caution when encountering" & _
           " this person as his sphere of influince" & _
           " is of a variety! (i.e. Charm, Wit)" & vbCrLf & _
           "This has been a broadcast of the" & _
           " National Blonde Moment Society...", vbExclamation, "" & _
           Apps_Title
    End
' Then end!
End Sub
