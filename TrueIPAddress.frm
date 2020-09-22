VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIpAddress 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "True IP - by Airr"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000C0&
   Icon            =   "TrueIPAddress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCheckExtIp 
      Caption         =   "Check"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCheckLocIp 
      Caption         =   "Check"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   960
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCopyExtAddress 
      Caption         =   "Copy"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopyLocAddress 
      Caption         =   "Copy"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Restart"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblLocIpAddress 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Press Check Button"
      Height          =   420
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblExtIpAddress 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Press Check Button"
      Height          =   420
      Left            =   3480
      TabIndex        =   6
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblExtAddress 
      Alignment       =   2  'Center
      Caption         =   "External IP Address"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblLocAddress 
      Alignment       =   2  'Center
      Caption         =   "Local IP Address"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmIpAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdCheckExtIp_Click()

    'check for connection error
    On Error GoTo testErr

    'set up variables
    Dim site As String, ip As Integer, delim As Integer

    'Submit query for IP via external website, and assign result to "s"
    site = Inet1.OpenURL("http://www.whatismyip.com")
    
    ' should trap these for 0 return value on the Instr functions!
    
    ip = InStr(site, "Your IP is ") ' find the ip address location on page
    delim = InStr(ip, site, "<br>") ' find the break after the number
    lblExtIpAddress = (Mid(site, delim - 15, 15)) ' get up to 15 characters comprising IP address before the break

    Exit Sub

testErr:
    If Err.Number = 5 Then
        lblExtIpAddress = "No Connection"
    End If

End Sub


Private Sub cmdCheckLocIp_Click()
        
    'Query machine for it's Local IP Address
    lblLocIpAddress = Winsock1.LocalIP

End Sub

Private Sub cmdCopyExtAddress_Click()
Clipboard.Clear
Clipboard.SetText lblExtIpAddress

End Sub

Private Sub cmdCopyLocAddress_Click()
Clipboard.Clear
Clipboard.SetText lblLocIpAddress
End Sub

Private Sub cmdQuit_Click()
Select Case MsgBox("End TrueIP?", vbYesNo Or vbInformation, "Quit")

    Case vbYes: Unload Me
    Case Else:
    
End Select



End Sub

Private Sub cmdRefresh_Click()
    Call RestartMe
    
End Sub

