VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSensor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP/Host Name Sensor By Slippah"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
   Icon            =   "frmSensor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   3915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopyHostName 
      Caption         =   "Copy Host Name"
      Height          =   400
      Left            =   2520
      TabIndex        =   5
      Top             =   580
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Your Host Name:"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   500
      Width           =   2295
      Begin VB.Label HostName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   240
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdCopyIP 
      Caption         =   "Copy IP Address"
      Height          =   400
      Left            =   2520
      TabIndex        =   1
      Top             =   75
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Your IP Address:"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.Label IPAddress 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   2115
      End
   End
   Begin MSWinsockLib.Winsock IPSensor 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSensor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCopyIP_Click()
Clipboard.Clear
Clipboard.SetText IPAddress.Caption
End Sub

Private Sub cmdCopyHostName_Click()
Clipboard.Clear
Clipboard.SetText HostName.Caption
End Sub

Private Sub Form_Load()
IPAddress.Caption = IPSensor.LocalIP
HostName.Caption = IPSensor.LocalHostName
End Sub
Private Sub IPSensor_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Your IP address or your host name cannot be determined at this time.", vbCritical, "Error"
End Sub
