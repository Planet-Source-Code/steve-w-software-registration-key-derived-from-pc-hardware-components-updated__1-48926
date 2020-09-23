VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Registration Demo"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "3. User registers app to unlock full version features"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   4935
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register"
         Height          =   285
         Left            =   3600
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtGetRegKey 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Text            =   "Paste Reg Key here"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtValidationStatus 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Text            =   "ValidationStatus"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Registration Status"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   750
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "2. Author generates reg key && sends to user"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   4935
      Begin VB.CommandButton cmdCopyRegKey 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtGetId 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Text            =   "Paste Machine ID here"
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtRegKey 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Text            =   "RegKey"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Registration Key"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   750
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "1. User sends machine specific ID to software author"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      Begin VB.TextBox txtId 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Text            =   "MachineId"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdCopyId 
         Caption         =   "Copy"
         Height          =   285
         Left            =   3600
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Machine ID"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   390
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code shows how to register an app to individual PCs using unique properties such as the CPU ID,OS serial, MAC address etc.
'It uses WMI so if you're running W95, W98 or NT4 & you get an automation error download & install the WMI engine from: http://download.microsoft.com/download/platformsdk/wmicore/1.5/W9XNT4/EN-US/wmicore.EXE.
'The app generates an ID code that is derived from properties unique to the user's PC - this example uses the CPU ID & OS serial number, but this can easily be adapted to use an NIC mac address and/or pretty much any other hardware device property.
'A user wishing to register the app would send the ID code to the software author who uses it to generate a registration key (which is based on the same hardware properties as the ID code). This is sent to the user enabling registration. Once the app has been registered the machine ID serves no purpose & the user can re-register the app at any time with the same registration key provided the components used to generate the original ID code have not been changed. In the event of a hardware change the app will have to be re-registered to run although it would be simple to implement a system similar to MS WPA allowing some changes.
'There is no need to hide the ID code & registration key so they can be stored as plain text in an ini file (as in this example) or in the registry.
'This method isn't foolproof but it can be used as the basis of a registration system that will keep most people honest. The ID gen, key gen & validator are all contained in the demo for simplicity but the key gen would of course be separate in a real app.
'The code is simple & has lots of comments but email steve@cantec.net if you have any problems.
'Feel free to modify & use it as you see fit & please vote if you find it useful. Any feedback or comments are welcome.

Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long 'For ini read
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long 'For ini write
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Dim varMachineId As String
Dim ValidateId As Boolean
Dim RegistrationSucceeded As Boolean
Dim varKey As String
Dim varRegistryPath As String

Private Sub GetMachineInfo()
  Dim SWbemSet(2) As SWbemObjectSet 'Need to incl project reference 'Microsoft WMI Scripting Library'
  Dim SWbemObj As SWbemObject
  Dim varObjectToId(2) As String
  Dim varSerial(2) As String
  Dim i, j As Integer
  Dim varSerialHex As String
  Dim varSerialTemp(2) As String
  varObjectToId(1) = "Win32_Processor,ProcessorId"
  varObjectToId(2) = "Win32_OperatingSystem,SerialNumber"
  'We're using CPU & OS serials but can be any WMI class object property that returns a numeric value (or alpha if you use an algorithm to convert characters to numbers - LSB of asc(character) maybe...)
  'Refer to http://msdn.microsoft.com/library/en-us/wmisdk/wmi/retrieving_a_class.asp for more info
  For i = 1 To 2 'we're only using 2 objects in this example - can be more but code will have to me modified to suit
    Set SWbemSet(i) = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf(Split(varObjectToId(i), ",")(0))
    varSerial(i) = ""
    For Each SWbemObj In SWbemSet(i) 'Is buggy if querying 2nd similar device (eg 2nd CPU ID or MAC Address) but I ran out of time
      varSerialHex = ""
      varSerialTemp(i) = ""
      varSerial(i) = SWbemObj.Properties_(Split(varObjectToId(i), ",")(1)) 'Property value
      For j = 1 To Len(varSerial(i)) 'Strip out any non hex characters so we can do some simple maths later
        varSerialHex = Mid(varSerial(i), j, 1)
        If varSerialHex Like "[0-9A-Fa-f]" Then varSerialTemp(i) = varSerialTemp(i) & varSerialHex
      Next
    Next
    varSerial(i) = varSerialTemp(i)
    varSerial(i) = Right(varSerial(i), 4) 'Let's just use last 4 digits of each serial for simplicity
  Next
  varMachineId = ""
  For i = 1 To 4
    varMachineId = varMachineId & Mid(varSerial(1), i, 1) & Mid(varSerial(2), i, 1) 'Just a little obfuscation
  Next
End Sub

Private Sub cmdCopyId_Click()
  Clipboard.SetText (txtId.Text)
  txtGetId.Visible = True
End Sub

Private Sub cmdCopyRegKey_Click()
  Clipboard.SetText (txtRegKey.Text)
  txtGetRegKey.Visible = True
End Sub

Private Sub cmdRegister_Click()
  If RegistrationSucceeded = False Then
    Dim varRegKey As String
    Dim varRandNum(2) As String
    Dim varRandNumPosn(2) As Integer
    Dim i As Integer
    varRegKey = Replace(txtGetRegKey.Text, "-", "")
    txtValidationStatus.Text = "Processing..."
    txtValidationStatus.Refresh
      Call WriteINI("SETTINGS", "RegistrationKey", varRegKey, App.Path & "\settings.ini")
      ValidateRegistration
    If RegistrationSucceeded = False Then
      MsgBox ("Shame on you! You have attempted to use an invalid registration key."), vbCritical
      Call WriteINI("SETTINGS", "RegistrationKey", "", App.Path & "\settings.ini")
      'End
    End If
  Else
    Call WriteINI("SETTINGS", "MachineId", "", App.Path & "\settings.ini")
    Call WriteINI("SETTINGS", "RegistrationKey", "", App.Path & "\settings.ini")
    Call Form_Load
  End If
End Sub

Private Sub Form_Load()
Call HackerScan
Dim i, j, tmp As Integer
Dim varKeyTmp(16) As String
Dim varKey2 As String
Dim varKeyOut As String
  varKey = "320420352196268181894210437267790267909582698624445212800344381963372357131848839301613255938768148682177489084666712402256650210971"
  varKeyTmp(0) = varKey
  varKey2 = Mid(varKey, 24, 16)
  varKey2 = Replace(varKey2, "0", "1")
  For i = 1 To 16
    varKeyTmp(i) = Mid(varKey, Mid(varKey2, i, 1), 120)
  Next
For i = 1 To 120
  For j = 0 To 15
    tmp = Val(Mid(varKeyTmp(j), i, 1)) + Val(Mid(varKeyTmp(j + 1), i, 1))
    If tmp > 9 Then tmp = tmp - 10
  Next
  varKeyOut = varKeyOut & tmp
Next

For i = 1 To 16
  varKeyTmp(i) = ""
Next

  varKey = ""
  varKey = varKeyOut
  'MsgBox (varKey): End
  varMachineId = ReadINI("SETTINGS", "MachineId", App.Path & "\settings.ini") 'All ini references would be stored in the registry but I didn't want to bother with the extra API in this example
  If varMachineId = "" Then Call CreateMachineId 'Let's create a Machine ID if one doesn't already exist
  txtId.Text = varMachineId
  txtGetId = "Paste Machine ID here"
  txtGetRegKey = "Paste Reg Key here"
  txtGetId.Visible = False
  'txtGetRegKey.Visible = False
  cmdCopyRegKey.Enabled = False
  ValidateRegistration
  If RegistrationSucceeded = False Then MsgBox ("You are currently running in UNREGISTERED mode, some functionality is disabled"), vbInformation
End Sub

Private Sub CreateMachineId()
  Call HackerScan
  Dim varRandNum(2) As String
  Dim varRandNumPosn(2) As Integer
  Dim i As Integer
  GetMachineInfo
  varMachineId = CLng("&H" & varMachineId)
  Randomize
  For i = 1 To 2
    varRandNumPosn(i) = CInt(Int(90 * Rnd() + 10)) 'all we're doing here is getting a couple of random but repeatable numbers - bounds are set to get a 2 digit start position
    varRandNum(i) = Mid(varKey, varRandNumPosn(i), 8) 'get an 8 digit number string from the varKey string starting at position varRandNumPosn
  Next
  varMachineId = Str(Val(varMachineId) - Val(varRandNum(1)) + Val(varRandNum(2))) 'You could make this much more complicated but this is only a simple example
  varMachineId = varMachineId & varRandNumPosn(1) & varRandNumPosn(2) 'by storing the start position in the machine ID itself we can get the actual random number back later
  varMachineId = Trim(varMachineId) 'strip the leading space that can occur
  Call WriteINI("SETTINGS", "MachineId", varMachineId, App.Path & "\settings.ini")
End Sub

Private Sub ValidateRegistration()
  Call HackerScan
  Dim varRandNum(2) As String
  Dim varRandNumPosn(2) As Integer
  Dim i As Integer
  Dim varRegKey As String
  Dim varIdLookup As String
  varRegKey = ReadINI("SETTINGS", "RegistrationKey", App.Path & "\settings.ini")
  If varRegKey = "" Then
    txtValidationStatus.Text = "UNREGISTERED"
    GoTo registration 'Yeah - I know
  End If
  varRandNumPosn(1) = Mid(varRegKey, 1, 2) 'All this stuff just extracts the 8 digit hex number that will match the component IDs (ProcessorID & OS SerialNumber) value - work thru it & it'll be clear
  varRandNumPosn(2) = Mid(varRegKey, 3, 2)
  varRegKey = Right(varRegKey, Len(varRegKey) - 4)
  varRandNum(1) = Mid(varKey, varRandNumPosn(1), 11)
  varRandNum(2) = Mid(varKey, varRandNumPosn(2), 11)
  varRegKey = Str(Val(varRegKey) + Val(varRandNum(1)) - Val(varRandNum(2)))
  varRandNumPosn(1) = Mid(varRegKey, Len(varRegKey) - 3, 2)
  varRandNumPosn(2) = Mid(varRegKey, Len(varRegKey) - 1, 2)
  For i = 1 To 2
    varRandNum(i) = Mid(varKey, varRandNumPosn(i), 8)
  Next
  varRegKey = Mid(varRegKey, 1, Len(varRegKey) - 4)
  varRegKey = Val(varRegKey) - Val(varRandNum(2)) + Val(varRandNum(1))
  varRegKey = Hex(varRegKey)
  If Len(varRegKey) < 8 Then 'Pad with leading '0's if necessary to get 8 digits
    Do Until Len(varRegKey) = 8
      varRegKey = "0" & varRegKey
    Loop
  End If
  GetMachineInfo 'Get the actual 8 digit hex number that represents the component IDs (ProcessorID & OS SerialNumber)
registration:
  If varRegKey = varMachineId Then RegistrationSucceeded = True Else RegistrationSucceeded = False 'If they match we have a valid registration number
  If RegistrationSucceeded = True Then
    'remove 'unregistered version' restrictions here
    cmdRegister.Caption = "Unregister" 'Just some stuff for the example
    txtValidationStatus.Text = "REGISTERED"
    cmdCopyId.Enabled = False
    cmdCopyRegKey.Enabled = False
    cmdRegister.Enabled = True
  Else
    'Run as unregistered version
    cmdRegister.Caption = "Register" 'Just some stuff for the example
    txtValidationStatus.Text = "UNREGISTERED"
    cmdCopyId.Enabled = True
  End If
  txtValidationStatus.Refresh
  Form1.Caption = "Registration Demo - " & txtValidationStatus.Text
  Form1.Refresh
End Sub

Private Sub txtGetId_Change() 'This is the key generator
  ValidateMachineId
  Dim varRandNum(2) As String
  Dim varRandNumPosn(2) As Integer
  Dim i As Integer
  Dim varRegKey As String
  If txtGetId.Text = "" Then Exit Sub
  Randomize
  For i = 1 To 2 'This should be looking familiar by now
    varRandNumPosn(i) = CInt(Int(90 * Rnd() + 10))
    varRandNum(i) = Mid(varKey, varRandNumPosn(i), 11)
  Next
  varRegKey = Str(Val(txtGetId.Text) - Val(varRandNum(1)) + Val(varRandNum(2))) 'Use your imagination - this can be as complicated as you like in a real app of course...
  varRegKey = varRandNumPosn(1) & varRandNumPosn(2) & Replace(varRegKey, " ", "") 'Store the random start positions in the key for decoding
  txtRegKey.Text = varRegKey
  For i = Len(txtRegKey.Text) To 1 Step -1 'Don't need to do this but it makes the key a bit easier for the user to read when registering
    Select Case i
      Case 5, 10, 15
        txtRegKey.SelStart = i: txtRegKey.SelLength = 0
        If i < Len(txtRegKey.Text) Then txtRegKey.SelText = txtRegKey.SelText & "-"
    End Select
  Next
  cmdCopyRegKey.Enabled = True
End Sub

Private Sub txtGetId_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  txtGetId.SelStart = 0 'Just making it easy for you!
  txtGetId.SelLength = Len(txtGetId.Text)
End Sub

Private Sub txtGetRegKey_Change()
  cmdRegister.Enabled = True
End Sub

Private Sub txtGetRegKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  txtGetRegKey.SelStart = 0
  txtGetRegKey.SelLength = Len(txtGetRegKey.Text)
End Sub
Private Sub ValidateMachineId()
  Call HackerScan
  Dim varRandNum(2) As String
  Dim varRandNumPosn(2) As Integer
  Dim i As Integer
  Dim varIdLookup As String
  Dim ValidateId As Boolean
  varIdLookup = txtId.Text
  varRandNumPosn(1) = Mid(varIdLookup, Len(varIdLookup) - 3, 2)
  varRandNumPosn(2) = Mid(varIdLookup, Len(varIdLookup) - 1, 2)
  For i = 1 To 2
    varRandNum(i) = Mid(varKey, varRandNumPosn(i), 8)
  Next
  varIdLookup = Mid(varIdLookup, 1, Len(varIdLookup) - 4)
  varIdLookup = Val(varIdLookup) - Val(varRandNum(2)) + Val(varRandNum(1))
  varIdLookup = Hex(varIdLookup)
  If Len(varIdLookup) < 8 Then 'Pad with leading '0's if necessary to get 8 digits
    Do Until Len(varIdLookup) = 8
      varIdLookup = "0" & varIdLookup
    Loop
  End If
  GetMachineInfo
  If varIdLookup = varMachineId Then ValidateId = True Else ValidateId = False
  If ValidateId = False Then
    MsgBox ("You must provide a valid Machine ID in order to register this product." & Chr(13) & "The application will now exit."), vbCritical
    CreateMachineId 'Make a new ID otherwise they won't ever be able to register
    End
  End If
End Sub

Public Function ReadINI(strsection As String, strkey As String, strfullpath As String) As String
  Dim strBuffer As String
  strBuffer$ = String$(32, Chr$(0&))
  ReadINI$ = Left$(strBuffer$, GetPrivateProfileString(strsection$, ByVal LCase$(strkey$), "", strBuffer, Len(strBuffer), strfullpath$))
End Function

Public Function WriteINI(strsection As String, strkey As String, strkeyvalue As String, strfullpath As String)
  Call WritePrivateProfileString(strsection$, UCase$(strkey$), strkeyvalue$, strfullpath$)
End Function

Public Sub HackerScan() 'This sub is an adaptation of Kevin Lingofelter's code at http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=10000&lngWId=1
  Dim hFile As Long, retVal As Long
  Dim clsMonitor As String
  
  clsMonitor = "1" & "8" & "4" & "6" & "7" & "-" & "4" & "1" 'This checks for latest versions of RegMon & FileMon only, classes of other 'hacking' apps can be added here - classname is split to hide from hex editors
  hFile = CreateFile("\\.\SICE", GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0) ' SoftICE (W9x)is detected.
  If hFile <> -1 Then hFile = CreateFile("\\.\NTICE", GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0) ' SoftICE (NT)is detected.
  If FindWindow(clsMonitor, vbNullString) <> 0 Or hFile <> -1 Then MsgBox ("Hacking activity detected! Application will exit"): End
  If hFile <> -1 Then retVal = CloseHandle(hFile)
End Sub
