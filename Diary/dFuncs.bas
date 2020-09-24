Attribute VB_Name = "dFuncs"
Option Explicit
' Ini Api Calls.
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'm_file is the file we operate under
'm_buffer is the number of characters to retrieve max
'   -- Need to set this high (over 5000) if you plan on
'      Read_Sections or Read_Keys a large INI
Dim m_File As String, m_Buffer As Long
' Used to set the shape of the form
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
' Used to create the rounded rectangle region
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' Used to make the form draggable
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' Also used to make the form draggable
Public Declare Function ReleaseCapture Lib "user32" () As Long
' Used to make the window always on top
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Attribute SetWindowPos.VB_MemberFlags = "10"
' Various constants used by the above functions
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
' Make roundness of the form.
Public Sub DoTransparency(TheForm As Form)
' TheForm:  The form you want to be rounded rectangle shape
    Dim FormWidthInPixels As Long
    Dim FormHeightInPixels As Long
    Dim a
' Convert the form's height and width from twips to pixels
    FormWidthInPixels = TheForm.Width / Screen.TwipsPerPixelX
    FormHeightInPixels = TheForm.Height / Screen.TwipsPerPixelY
' Make a rounded rectangle shaped region with the dimentions of the form
    a = CreateRoundRectRgn(0, 0, FormWidthInPixels, FormHeightInPixels, 15, 15)
' Set this region as the shape for "TheForm"
    a = SetWindowRgn(TheForm.hWnd, a, True)
End Sub
' Set the window on top or notontop.
Public Sub AlwaysOnTop(TheForm As Form, Toggle As Boolean)
' TheForm:  The form you want to make always on top or not
' Toggle:   Boolean (True/False) - True for always on top, False for normal
    If Toggle = True Then
        SetWindowPos TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    Else
        SetWindowPos TheForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    End If
End Sub
' Allow dragging of application from title bar.
Public Sub DoDrag(TheForm As Form)
' TheForm:  The form you want to start dragging
    ReleaseCapture
    SendMessage TheForm.hWnd, &HA1, 2, 0&
End Sub
' Logg lamers trying to enter application.
Public Sub DoLogg()
Attribute DoLogg.VB_MemberFlags = "10"
' Contains the number of times accessed.
Dim lamerVal As Integer
' Setup the iniFunction to point at the iniFile.
INISetup App.Path & "\System.ini", 10000
' Read times accessed allready.
lamerVal = Int(Trim(Read_Ini("System", "/*-*"))) + 1
' Write new accesses.
Write_Ini "System", "/*-*", lamerVal
' Log the date and time accessed.
Write_Ini "System", "*d", CStr(Date) & " " & CStr(Time)
End Sub
' Get all log info and show the user.
Public Sub Getlogg()
' Setup ini functions.
INISetup App.Path & "\System.ini", 10000
' Stores amount of times users tried to access.
Dim lTimes As String
' Stores Date and Time users tried to access.
Dim sDateTime As String
' Read the num of times into lTimes.
lTimes = Read_Ini("System", "/*-*")
' Read the date and time into sDateTime.
sDateTime = Read_Ini("System", "*d")
' The form to read log file.
With frmLog
' The label which the times accessed will show in.
    .lUsers.Caption = "There has been " & lTimes & _
                    " User(s) that has accessed" & _
                    " this application without a" & _
                    " valid password."
' The label which the data and time accessed will show in.
    .lDate = sDateTime
End With
End Sub
' Get the Password from file.
Public Function GetPass(section As Integer) As String
Dim FileNum As Integer
Dim Password As String
Dim pNum As Integer
pNum = Int(PassLen)
' Get a free file.
FileNum = FreeFile
' Open up Diary.pwd for Binary Access.
 Open App.Path & "\Diary.pwd" For Binary As FileNum
' Password will store a max of 9 chars.
    Password = String(pNum, " ")
' Now get the Section=Byte location, Password=Stored pass.
        Get #FileNum, section, Password
' Close up the file.
    Close FileNum
' Set the password to GetPass for retreval.
 GetPass = Password
End Function
' Make or Change password.
Public Function DoPass(section As Integer, Password As String) As Boolean
' Incase an Error triger BigBandAid _
  so not to close application.
On Error GoTo BigBandAid
' Default pass "040705672"
Dim FileNum As Integer
' Get a FreeFile number
FileNum = FreeFile
' open diary.pwd for binary access.
Open App.Path & "\Diary.pwd" For Binary As FileNum
' Save Section=Byte location, Password=Passy.
Put #FileNum, section, Password
' Close up the file
Close FileNum
StorePassLen Len(Password)
' Signal flag for password stored.
DoPass = True
' Now exit the function so not to goto BigBandAid my mistake.
Exit Function
' Error handler.
BigBandAid:
' Msg to inform the user an error occured.
    MsgBox "Error! Or incomplete Password." & _
           " Please try again. If this problem" & _
           " Occurs again please seek help from a" & _
           " local physician.", vbExclamation, "" & _
           "Felicias' Diary"
End Function
' Read the length of the password.
Private Function PassLen() As String
' Set up the ini function to read min/max of 10000 chars.
INISetup App.Path & "\System.ini", 10000
' Store the data found in PassLen
PassLen = Read_Ini("System", Trim("/!*l*"))
End Function
' When user creates a new password store the length.
Public Function StorePassLen(eValue)
' This writes to the allready made ini Section and key
' to store the length of the new password.
Write_Ini "System", "/!*l*", eValue
End Function
' Set up the Ini functions.
' BufferSize is how many chars will be read back
Public Sub INISetup(FileName As String, BufferSize As Long)
Attribute INISetup.VB_MemberFlags = "10"
    m_Buffer = BufferSize
    m_File = FileName
End Sub
' Read the ini file.
Public Function Read_Ini(iSection As String, iKeyName As String, Optional iDefault As String)
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    Dim Ret As String, NC As Long
    'Create the buffer
    Ret = String(m_Buffer, 0)
    'Retrieve the string
    NC = GetPrivateProfileString(iSection, iKeyName, iDefault, Ret, m_Buffer, m_File)
    'NC is the number of characters copied to the buffer
    If NC <> 0 Then
        Ret = Left$(Ret, NC)
    Else
        'Make sure to cut it down to number of char's returned
        Ret = ""
    End If
    'Turn the funky vbcrlf string into VBCRLFs
    Ret = Replace(Ret, "%%&&Chr(13)&&%%", vbCrLf)
    'Return the setting
    Read_Ini = Ret
End Function
' Write to the ini file.
Public Sub Write_Ini(iSection As String, iKeyName As String, iValue As Variant)
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    'Make sure to change it to a String
    iValue = CStr(iValue)
    'Turn all vbcrlf's into that funky string
    iValue = Replace(iValue, vbCrLf, "%%&&Chr(13)&&%%")
    WritePrivateProfileString iSection, iKeyName, CStr(iValue), m_File
End Sub
' Read of what the first section is.
Public Function Read_Sections()
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    Dim Ret As String, NC As Long
    'Create the buffer
    Ret = String(m_Buffer, 0)
    'Retrieve the string, return '[-na-]' if there is none
    NC = GetPrivateProfileString(vbNullString, vbNullString, vbNullString, Ret, m_Buffer, m_File)
    'NC is the number of characters returned
    If NC <> 0 Then
        Ret = Left$(Ret, NC - 1)
    End If
    'Return the sections
    Read_Sections = Ret
End Function
' Read the first key name in the section.
Public Function Read_Keys(iSection As String)
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    Dim Ret As String, NC As Long
    'Create the buffer
    Ret = String(m_Buffer, 0)
    'Retrieve the string, return '[-na-]' if there is none
    NC = GetPrivateProfileString(iSection, vbNullString, vbNullString, Ret, m_Buffer, m_File)
    'NC is the number of characters copied to the buffer
    If NC <> 0 Then
        Ret = Left$(Ret, NC - 1)
    End If
    'Return the sections
    Read_Keys = Ret
End Function
' Delete section of the ini.
Public Function DeleteSection(iSection As String)
'Haven't tested these two myself =\
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    WritePrivateProfileString iSection, vbNullString, vbNullString, m_File
End Function
' Delete key from Section.
Function DeleteKey(iSection As String, iKeyName As String)
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    WritePrivateProfileString iSection, iKeyName, vbNullString, m_File
End Function
' Open a new Diary file or txt file.
Public Function FileOpen() As String
' Error handler.
On Error GoTo HadaBadDay
' String to hold Filters.
Dim Filter As String
' These are the two filters used in this app.
Filter = "Text Files (*.txt)|*.txt|"
Filter = Filter + "Diary Files (*.dry)|*.dry|"
' Add new filters to dialogs' filters.
frmMain.cd1.Filter = Filter
' Set the filter index to op *.dry files first.
frmMain.cd1.FilterIndex = 2
' Set the action to open files.
frmMain.cd1.Action = 1
' Save the path to FileOpen string
FileOpen = frmMain.cd1.FileName
' Create a new FreeFile number.
Dim FileNum As Integer
FileNum = FreeFile
' Open the file for input.
Open FileOpen For Input As FileNum
' Load each line in file while not end of file.
While Not EOF(FileNum)
' Set all loaded data to the text box on main form.
frmMain.Text1.Text = Input(LOF(FileNum), FileNum)
' loop it
Wend
' Close the door.
Close FileNum
' Exit function so we don't run into error handler.
Exit Function
' Error handler.
HadaBadDay:
' Msg to let user know he canceled the dialog box.
    MsgBox "You Canceled the dialog box!"
' Exit function.
Exit Function
End Function

Public Function SaveDfile()
' If error then goto error handler.
On Error GoTo NotEnoughSex
' Stored freefile number.
Dim FileNum As Integer
' Stored filters for commondialog.
Dim Filter As String
' Name of file to be saved.
Dim FileNames As String
' Store the freefile.
FileNum = FreeFile
' Save filters to filter.
Filter = "Diary Files (*.dry)|*.dry |"
' Add all filters to the dialog control.
frmMain.cd1.Filter = Filter
' Set the dialog to index 1 on open.
frmMain.cd1.FilterIndex = 1
' Action to open save dialog.
frmMain.cd1.Action = 2
' Store users filename.
FileNames = frmMain.cd1.FileName
' Open file to save new diary.
Open FileNames For Append As FileNum
' Save data from textbox.
Print #FileNum, frmMain.Text1.Text
' Clode the door.
Close FileNum
' Send a msg alerting user file has been saved.
MsgBox "Data has been saved!", vbExclamation, "Felicais Diary"
' Exit so not to run into Error Handler.
Exit Function
' Error handler.
NotEnoughSex:
' Msg user that file did not save due to error.
    MsgBox "File did not save!"
' Exit Function
Exit Function
End Function
