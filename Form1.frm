VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change the MenuShowDelay"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "QUIT"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change the menu show delay"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo error
a% = InputBox("Enter a number between 1 and 1000", _
"Start Menu Speed")
'a is the integer value of the input in the inputbox
'checking the input

If a% > 0 And a% < 1001 Then
    'input is a valid number between 1 and 1000
    'and a (integer) is to be converted in b (string)
    b$ = CStr(a%)

    'creating MenuShowDelay with it´s value
    '(if already exists it just changes the value)
    Call savestring("HKEY_CURRENT_USER", "Control Panel\Desktop", "MenuShowDelay", b$)

    'resetting computer
    MsgBox "Reset your Computer", , "Changes are made"
    t& = ExitWindowsEx(EWX_FORCE Or EWX_REBOOT, 0)
Else    'value is a number but not valid
    MsgBox "Not a valid number between 1 and 1000"
End If
Exit Sub

error:
    'error, input was not a valid number
    MsgBox "Invalid Data Input"
End Sub

Private Sub Command2_Click()
Unload Me
End
End Sub
